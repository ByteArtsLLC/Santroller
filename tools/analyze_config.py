#!/usr/bin/env python3
"""Export Santroller JSON profile settings to CSV or an Excel workbook."""

import argparse
import csv
import json
import math
import os
import re
import zipfile
from xml.etree import ElementTree

MAX_EXCEL_ROWS = 1_048_576
INVALID_SHEET_CHARACTERS = re.compile(r"[\[\]:*?/\\]")
XML_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"

ElementTree.register_namespace("", XML_NAMESPACE)
ElementTree.register_namespace("r", RELATIONSHIP_NAMESPACE)


def flatten_settings(value, path=""):
    """Yield every scalar setting as a path/value pair, retaining array indices."""
    if isinstance(value, dict):
        if not value:
            yield path or "(root)", "{}"
        else:
            for key, child in value.items():
                child_path = f"{path}.{key}" if path else str(key)
                yield from flatten_settings(child, child_path)
    elif isinstance(value, list):
        if not value:
            yield path or "(root)", "[]"
        else:
            for index, child in enumerate(value):
                yield from flatten_settings(child, f"{path}[{index}]")
    else:
        yield path or "(root)", value


def safe_sheet_name(name, used_names):
    """Return a valid, unique Excel worksheet name."""
    base_name = INVALID_SHEET_CHARACTERS.sub("_", name).strip("'") or "Profile"
    base_name = base_name[:31]
    candidate = base_name
    suffix_number = 2
    while candidate.casefold() in used_names:
        suffix = f" ({suffix_number})"
        candidate = f"{base_name[:31 - len(suffix)]}{suffix}"
        suffix_number += 1
    used_names.add(candidate.casefold())
    return candidate


def column_name(column_number):
    """Convert a one-based column number to its Excel column name."""
    result = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xml_cell(reference, value, style=None):
    attributes = {"r": reference}
    if style is not None:
        attributes["s"] = str(style)

    if isinstance(value, bool):
        attributes["t"] = "b"
        cell_value = "1" if value else "0"
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and not math.isfinite(value):
            attributes["t"] = "inlineStr"
            inline_string = ElementTree.Element(f"{{{XML_NAMESPACE}}}is")
            text = ElementTree.SubElement(
                inline_string, f"{{{XML_NAMESPACE}}}t"
            )
            text.text = str(value)
            cell = ElementTree.Element(f"{{{XML_NAMESPACE}}}c", attributes)
            cell.append(inline_string)
            return cell
        cell_value = str(value)
    else:
        attributes["t"] = "inlineStr"
        inline_string = ElementTree.Element(f"{{{XML_NAMESPACE}}}is")
        text = ElementTree.SubElement(inline_string, f"{{{XML_NAMESPACE}}}t")
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = "null" if value is None else str(value)
        cell = ElementTree.Element(f"{{{XML_NAMESPACE}}}c", attributes)
        cell.append(inline_string)
        return cell

    cell = ElementTree.Element(f"{{{XML_NAMESPACE}}}c", attributes)
    ElementTree.SubElement(cell, f"{{{XML_NAMESPACE}}}v").text = cell_value
    return cell


def worksheet_xml(rows, widths, filter_columns=None):
    worksheet = ElementTree.Element(f"{{{XML_NAMESPACE}}}worksheet")
    last_column = column_name(len(widths))
    last_row = max(1, len(rows))
    ElementTree.SubElement(
        worksheet,
        f"{{{XML_NAMESPACE}}}dimension",
        {"ref": f"A1:{last_column}{last_row}"},
    )

    sheet_views = ElementTree.SubElement(
        worksheet, f"{{{XML_NAMESPACE}}}sheetViews"
    )
    sheet_view = ElementTree.SubElement(
        sheet_views,
        f"{{{XML_NAMESPACE}}}sheetView",
        {"workbookViewId": "0"},
    )
    ElementTree.SubElement(
        sheet_view,
        f"{{{XML_NAMESPACE}}}pane",
        {
            "ySplit": "1",
            "topLeftCell": "A2",
            "activePane": "bottomLeft",
            "state": "frozen",
        },
    )
    ElementTree.SubElement(worksheet, f"{{{XML_NAMESPACE}}}sheetFormatPr", {"defaultRowHeight": "18"})

    columns = ElementTree.SubElement(worksheet, f"{{{XML_NAMESPACE}}}cols")
    for index, width in enumerate(widths, 1):
        ElementTree.SubElement(
            columns,
            f"{{{XML_NAMESPACE}}}col",
            {"min": str(index), "max": str(index), "width": str(width), "customWidth": "1"},
        )

    sheet_data = ElementTree.SubElement(worksheet, f"{{{XML_NAMESPACE}}}sheetData")
    for row_index, values in enumerate(rows, 1):
        row = ElementTree.SubElement(
            sheet_data, f"{{{XML_NAMESPACE}}}row", {"r": str(row_index)}
        )
        for column_index, value in enumerate(values, 1):
            row.append(
                xml_cell(
                    f"{column_name(column_index)}{row_index}",
                    value,
                    style=1 if row_index == 1 else None,
                )
            )

    if filter_columns:
        ElementTree.SubElement(
            worksheet,
            f"{{{XML_NAMESPACE}}}autoFilter",
            {"ref": f"A1:{column_name(filter_columns)}{last_row}"},
        )

    ElementTree.SubElement(worksheet, f"{{{XML_NAMESPACE}}}pageMargins", {
        "left": "0.7",
        "right": "0.7",
        "top": "0.75",
        "bottom": "0.75",
        "header": "0.3",
        "footer": "0.3",
    })
    return ElementTree.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def workbook_parts(sheet_names):
    content_types = ElementTree.Element(
        f"{{{CONTENT_TYPES_NAMESPACE}}}Types"
    )
    ElementTree.SubElement(
        content_types,
        f"{{{CONTENT_TYPES_NAMESPACE}}}Default",
        {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"},
    )
    ElementTree.SubElement(
        content_types,
        f"{{{CONTENT_TYPES_NAMESPACE}}}Default",
        {"Extension": "xml", "ContentType": "application/xml"},
    )
    for part_name, content_type in (
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
    ):
        ElementTree.SubElement(
            content_types,
            f"{{{CONTENT_TYPES_NAMESPACE}}}Override",
            {"PartName": part_name, "ContentType": content_type},
        )
    for index in range(1, len(sheet_names) + 1):
        ElementTree.SubElement(
            content_types,
            f"{{{CONTENT_TYPES_NAMESPACE}}}Override",
            {
                "PartName": f"/xl/worksheets/sheet{index}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            },
        )

    root_relationships = ElementTree.Element(
        f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationships"
    )
    ElementTree.SubElement(
        root_relationships,
        f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship",
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "xl/workbook.xml",
        },
    )

    workbook = ElementTree.Element(f"{{{XML_NAMESPACE}}}workbook")
    sheets = ElementTree.SubElement(workbook, f"{{{XML_NAMESPACE}}}sheets")
    workbook_relationships = ElementTree.Element(
        f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationships"
    )
    for index, name in enumerate(sheet_names, 1):
        ElementTree.SubElement(
            sheets,
            f"{{{XML_NAMESPACE}}}sheet",
            {
                "name": name,
                "sheetId": str(index),
                f"{{{RELATIONSHIP_NAMESPACE}}}id": f"rId{index}",
            },
        )
        ElementTree.SubElement(
            workbook_relationships,
            f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship",
            {
                "Id": f"rId{index}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": f"worksheets/sheet{index}.xml",
            },
        )
    ElementTree.SubElement(
        workbook_relationships,
        f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship",
        {
            "Id": f"rId{len(sheet_names) + 1}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "Target": "styles.xml",
        },
    )

    styles = ElementTree.fromstring(
        b"""<?xml version="1.0" encoding="UTF-8"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="2">
            <font><sz val="11"/><name val="Calibri"/></font>
            <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <fills count="3">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FF24527A"/><bgColor indexed="64"/></patternFill></fill>
          </fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
          </cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>"""
    )
    return {
        "[Content_Types].xml": ElementTree.tostring(
            content_types, encoding="utf-8", xml_declaration=True
        ),
        "_rels/.rels": ElementTree.tostring(
            root_relationships, encoding="utf-8", xml_declaration=True
        ),
        "xl/workbook.xml": ElementTree.tostring(
            workbook, encoding="utf-8", xml_declaration=True
        ),
        "xl/_rels/workbook.xml.rels": ElementTree.tostring(
            workbook_relationships, encoding="utf-8", xml_declaration=True
        ),
        "xl/styles.xml": ElementTree.tostring(
            styles, encoding="utf-8", xml_declaration=True
        ),
    }


def load_config(input_path):
    with open(input_path, "r", encoding="utf-8") as config_file:
        data = json.load(config_file, parse_constant=reject_nonstandard_constant)

    if not isinstance(data, dict) or not isinstance(data.get("config"), dict):
        raise ValueError("JSON must contain a 'config' object")
    profiles = data["config"].get("profiles")
    if not isinstance(profiles, list):
        raise ValueError("'config.profiles' must be a list")
    for index, profile in enumerate(profiles, 1):
        if not isinstance(profile, dict):
            raise ValueError(f"Profile {index} must be a JSON object")
    return data, profiles


def reject_nonstandard_constant(value):
    raise ValueError(f"Invalid JSON number: {value}")


def csv_value(value):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return str(value).lower()
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def create_csv(data, profiles, output_path):
    with open(output_path, "w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.writer(output_file)
        writer.writerow(["Profile", "Setting", "Value"])

        for index, profile in enumerate(profiles, 1):
            options = profile.get("opts")
            profile_name = options.get("name") if isinstance(options, dict) else None
            if not isinstance(profile_name, str) or not profile_name.strip():
                profile_name = f"Profile {index}"
            for setting, value in flatten_settings(profile):
                writer.writerow([csv_value(profile_name), csv_value(setting), csv_value(value)])

        global_config = {
            key: value for key, value in data["config"].items() if key != "profiles"
        }
        global_settings = {
            **{key: value for key, value in data.items() if key != "config"},
            "config": global_config,
        }
        for setting, value in flatten_settings(global_settings):
            writer.writerow(["Global config", csv_value(setting), csv_value(value)])


def create_workbook(data, profiles, output_path):
    overview_rows = [[
        "Profile",
        "UID",
        "Device to emulate",
        "Mappings",
        "LEDs",
        "Assignments",
        "Standby wake",
        "Settings",
        "Settings sheet",
    ]]
    sheets = []
    used_names = {"overview"}

    for index, profile in enumerate(profiles, 1):
        options = profile.get("opts")
        if not isinstance(options, dict):
            options = {}
        profile_name = options.get("name")
        if not isinstance(profile_name, str) or not profile_name.strip():
            profile_name = f"Profile {index}"

        sheet_name = safe_sheet_name(profile_name, used_names)
        settings = list(flatten_settings(profile))
        if len(settings) + 1 > MAX_EXCEL_ROWS:
            raise ValueError(
                f"Profile '{profile_name}' exceeds Excel's worksheet row limit"
            )

        overview_rows.append([
            profile_name,
            options.get("uid"),
            options.get("deviceToEmulate"),
            len(profile.get("mappings", [])) if isinstance(profile.get("mappings"), list) else 0,
            len(profile.get("leds", [])) if isinstance(profile.get("leds"), list) else 0,
            len(profile.get("assignments", [])) if isinstance(profile.get("assignments"), list) else 0,
            len(profile.get("standbyWake", [])) if isinstance(profile.get("standbyWake"), list) else 0,
            len(settings),
            sheet_name,
        ])
        rows = [["Setting", "Value"]]
        rows.extend(settings)
        sheets.append((sheet_name, rows, [72, 48], 2))

    global_config = {
        key: value for key, value in data["config"].items() if key != "profiles"
    }
    global_settings = list(flatten_settings({
        **{key: value for key, value in data.items() if key != "config"},
        "config": global_config,
    }))
    if len(global_settings) + 1 > MAX_EXCEL_ROWS:
        raise ValueError("Global config exceeds Excel's worksheet row limit")
    overview_name = "Overview"
    sheets.insert(
        0,
        (overview_name, overview_rows, [36, 12, 28, 12, 10, 14, 14, 12, 24], 9),
    )
    sheets.append((
        "Config Settings",
        [["Setting", "Value"], *global_settings],
        [72, 48],
        2,
    ))

    # Include config settings as a separate sheet without colliding with profile names.
    config_sheet = sheets.pop()
    config_name = safe_sheet_name(config_sheet[0], used_names)
    sheets.append((config_name, config_sheet[1], config_sheet[2], config_sheet[3]))
    if len(overview_rows) > MAX_EXCEL_ROWS:
        raise ValueError("There are too many profiles for the overview worksheet")

    parts = workbook_parts([sheet[0] for sheet in sheets])
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as workbook_file:
        for name, content in parts.items():
            workbook_file.writestr(name, content)
        for index, (_, rows, widths, filter_columns) in enumerate(sheets, 1):
            workbook_file.writestr(
                f"xl/worksheets/sheet{index}.xml",
                worksheet_xml(rows, widths, filter_columns),
            )


def main():
    parser = argparse.ArgumentParser(
        description="Summarize Santroller JSON config profiles in CSV or Excel."
    )
    parser.add_argument("input", help="Santroller JSON configuration file")
    parser.add_argument(
        "--format",
        choices=("csv", "xlsx"),
        default="csv",
        help="Output format (default: csv)",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output path (defaults to the input filename with the selected format's extension)",
    )
    args = parser.parse_args()

    output_path = args.output or os.path.splitext(args.input)[0] + f".{args.format}"
    if os.path.abspath(args.input) == os.path.abspath(output_path):
        parser.error("output must not overwrite the input JSON file")

    try:
        data, profiles = load_config(args.input)
        if args.format == "csv":
            create_csv(data, profiles, output_path)
        else:
            create_workbook(data, profiles, output_path)
    except (OSError, json.JSONDecodeError, ValueError) as error:
        parser.error(str(error))

    print(f"Created {output_path} ({len(profiles)} profiles, {args.format.upper()})")


if __name__ == "__main__":
    main()
