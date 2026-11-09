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
XML_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"

ElementTree.register_namespace("", XML_NAMESPACE)
ElementTree.register_namespace("r", RELATIONSHIP_NAMESPACE)


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


def profile_names(profiles):
    names = []
    seen = set()
    for index, profile in enumerate(profiles, 1):
        options = profile.get("opts")
        name = options.get("name") if isinstance(options, dict) else None
        if not isinstance(name, str) or not name.strip():
            name = f"Profile {index}"
        candidate = name
        suffix_number = 2
        while candidate.casefold() in seen:
            candidate = f"{name} ({suffix_number})"
            suffix_number += 1
        seen.add(candidate.casefold())
        names.append(candidate)
    return names


def humanize(value):
    value = str(value).replace("_", " ")
    value = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", value)
    value = value.strip().title()
    for abbreviation, replacement in (
        ("Gpio", "GPIO"),
        ("Usb", "USB"),
        ("Ws2812", "WS2812"),
        ("I2C", "I2C"),
        ("Rgb", "RGB"),
    ):
        value = value.replace(abbreviation, replacement)
    return value


def device_labels(config):
    labels = {}
    for device in config.get("devices", []):
        if not isinstance(device, dict):
            continue
        device_id = device.get("deviceid", device.get("deviceId"))
        device_type = next((key for key in device if key not in ("deviceid", "deviceId")), None)
        if device_id is not None and device_type:
            labels[device_id] = f"{humanize(device_type)} device {device_id}"
    return labels


def find_input_source(value):
    source_types = {
        "gpio", "matrix", "usb", "encoder", "i2c", "midi", "cycle",
        "accelerometer", "touch", "analog",
    }
    if isinstance(value, dict):
        for key, child in value.items():
            if key in source_types and isinstance(child, dict):
                if key == "cycle":
                    nested_source = find_input_source(child.get("input", child))
                    if nested_source:
                        return nested_source
                    continue
                return key, child
        for child in value.values():
            source = find_input_source(child)
            if source:
                return source
    elif isinstance(value, list):
        for child in value:
            source = find_input_source(child)
            if source:
                return source
    return None


def describe_input(value, labels):
    source = find_input_source(value)
    if not source:
        return "Other", "Always" if isinstance(value, dict) and "staticMapping" in value else "Unspecified"

    source_type, details = source
    device_id = details.get("deviceid", details.get("deviceId"))
    device = labels.get(device_id, f"{humanize(source_type)} device {device_id}") if device_id is not None else humanize(source_type)
    if source_type == "gpio":
        pin = details.get("pin")
        description = f"Pin {pin}" if pin is not None and pin >= 0 else "Unassigned pin"
    elif source_type == "matrix":
        pin = details.get("pin")
        output_pin = details.get("outputPin")
        parts = []
        if pin is not None:
            parts.append(f"Input {pin}")
        if output_pin is not None:
            parts.append(f"Output {output_pin}")
        description = ", ".join(parts) or "Unspecified input"
    else:
        description = ", ".join(
            f"{humanize(key)} {item}"
            for key, item in details.items()
            if key in ("pin", "button", "channel", "port")
        ) or "Input"
    return device, description


def mapping_label(mapping):
    if not isinstance(mapping, dict) or not mapping:
        return None, None
    key, value = next(iter(mapping.items()))
    labels = {
        "gamepadButton": "Gamepad button",
        "rbButton": "Rock Band button",
        "guitarHeroButton": "Guitar Hero button",
        "drumButton": "Drum button",
        "keyboardKey": "Keyboard key",
        "mouseButton": "Mouse button",
        "midiNote": "MIDI note",
    }
    if isinstance(value, (str, int, float, bool)):
        return labels.get(key, humanize(key)), humanize(value)
    return labels.get(key, humanize(key)), json.dumps(value, ensure_ascii=False, sort_keys=True)


def color_description(device):
    def color(prefix):
        red, green, blue = (
            device.get(f"{prefix}{channel}", 0) for channel in "RGB"
        )
        return f"#{red:02X}{green:02X}{blue:02X}"

    pressed = color("end")
    released = color("start") if device.get("hasStart") else "#000000"
    return f"Pressed: {pressed}; Released: {released}"


def comparison_table(config, profiles):
    names = profile_names(profiles)
    labels = device_labels(config)
    entries = {}

    def add(profile_index, device, usage, input_description, value, group_by_input=True):
        key = (device, usage, input_description if group_by_input else "")
        if key not in entries:
            entries[key] = {"inputs": [], "values": [[] for _ in profiles]}
        row = entries[key]
        if input_description not in row["inputs"]:
            row["inputs"].append(input_description)
        value = str(value)
        if value and value not in row["values"][profile_index]:
            row["values"][profile_index].append(value)

    for profile_index, profile in enumerate(profiles):
        for mapping in profile.get("mappings", []):
            if not isinstance(mapping, dict):
                continue
            usage, action = mapping_label(mapping.get("mapping"))
            if not usage:
                continue
            device, input_description = describe_input(mapping.get("input"), labels)
            add(profile_index, device, usage, input_description, action)

        for led in profile.get("leds", []):
            if not isinstance(led, dict):
                continue
            devices = led.get("device")
            if not isinstance(devices, dict) or not devices:
                continue
            device_type, device_settings = next(iter(devices.items()))
            if not isinstance(device_settings, dict):
                continue
            device_id = device_settings.get("deviceId", device_settings.get("deviceid"))
            device = labels.get(device_id, f"{humanize(device_type)} device {device_id}" if device_id is not None else humanize(device_type))
            input_device, input_description = describe_input(led.get("mapping"), labels)
            if input_description == "Always":
                input_description = "Always"
            elif input_device not in ("Other", device):
                input_description = f"{input_device}: {input_description}"

            if device_type == "rgb":
                led_indices = device_settings.get("activeLed", [])
                if not isinstance(led_indices, list) or not led_indices:
                    led_indices = ["All"]
                for led_index in led_indices:
                    usage = f"LED {led_index + 1}" if isinstance(led_index, int) else f"{led_index} LEDs"
                    behavior = color_description(device_settings)
                    if input_description != "Always":
                        behavior = f"{input_description}: {behavior}"
                    add(
                        profile_index,
                        device,
                        usage,
                        input_description,
                        behavior,
                        group_by_input=False,
                    )
            elif device_type == "display":
                element_type = humanize(device_settings.get("type", "Display element"))
                x, y = device_settings.get("x"), device_settings.get("y")
                position = f"Position {x}, {y}" if x is not None and y is not None else "Position unspecified"
                text = device_settings.get("text", device_settings.get("value", ""))
                add(profile_index, device, element_type, position, text or "(no text)")
            else:
                add(
                    profile_index,
                    device,
                    humanize(device_type),
                    input_description,
                    json.dumps(device_settings, ensure_ascii=False, sort_keys=True),
                )

        def find_input_assignments(value):
            if isinstance(value, dict):
                for key, child in value.items():
                    if key == "inputAnyTime" and isinstance(child, dict):
                        yield child
                    else:
                        yield from find_input_assignments(child)
            elif isinstance(value, list):
                for child in value:
                    yield from find_input_assignments(child)

        for assignment in find_input_assignments(profile.get("assignments", [])):
            assigned_input = assignment.get("input")
            cycle = assigned_input.get("cycle") if isinstance(assigned_input, dict) else None
            if isinstance(cycle, dict):
                usage = "Profile selection"
                cycle_id = cycle.get("deviceid", cycle.get("deviceId"))
                device = labels.get(cycle_id, f"Cycle device {cycle_id}")
                source = find_input_source(cycle.get("input"))
                input_device, source_description = describe_input(
                    cycle.get("input"), labels
                )
                if source and source[0] == "gpio" and source[1].get("pin", -1) < 0:
                    input_description = "Cycle control"
                    value_prefix = ""
                else:
                    input_description = (
                        f"{input_device}: {source_description}"
                        if input_device != "Other"
                        else source_description
                    )
                    value_prefix = f"{input_description}: "
                cycle_value = assignment.get("triggerValue")
                value = (
                    f"{value_prefix}Cycle value {cycle_value}"
                    if cycle_value is not None
                    else f"{value_prefix}Cycle control"
                )
                add(
                    profile_index,
                    device,
                    usage,
                    input_description,
                    value,
                    group_by_input=False,
                )
            else:
                device, input_description = describe_input(assigned_input, labels)
                usage = "Profile shortcut"
                value = humanize(assignment.get("trigger", "Assigned"))
                add(profile_index, device, usage, input_description, value)

        for wake_input in profile.get("standbyWake", []):
            if isinstance(wake_input, dict):
                device, input_description = describe_input(wake_input.get("input", wake_input), labels)
                add(profile_index, device, "Wake from standby", input_description, "Wake")

        options = profile.get("opts")
        if isinstance(options, dict):
            option_labels = {
                "deviceToEmulate": ("Controller", "Emulated device"),
                "faceButtonMappingMode": ("Controller", "Face-button layout"),
                "xinputOnWindows": ("Controller", "XInput on Windows"),
                "ps4OrPs5Mode": ("Controller", "PlayStation mode"),
                "supportsSlider": ("Controller", "Slider support"),
            }
            for option, (device, usage) in option_labels.items():
                if option in options:
                    add(profile_index, device, usage, "Profile option", humanize(options[option]))

    sort_key = lambda value: [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", value)
    ]
    rows = [["Device", "Usage", "Input", *names]]
    for (device, usage, input_key), row in sorted(
        entries.items(), key=lambda item: tuple(sort_key(part) for part in item[0])
    ):
        input_description = input_key or "; ".join(sorted(row["inputs"], key=sort_key))
        values = ["; ".join(profile_values) for profile_values in row["values"]]
        rows.append([device, usage, input_description, *values])
    return rows


def create_csv(config, profiles, output_path):
    rows = comparison_table(config, profiles)
    with open(output_path, "w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.writer(output_file)
        writer.writerows([[csv_value(value) for value in row] for row in rows])


def create_workbook(config, profiles, output_path):
    rows = comparison_table(config, profiles)
    if len(rows) > MAX_EXCEL_ROWS:
        raise ValueError("Comparison exceeds Excel's worksheet row limit")
    widths = [24, 28, 38, *([28] * len(profiles))]
    parts = workbook_parts(["Comparison"])
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as workbook_file:
        for name, content in parts.items():
            workbook_file.writestr(name, content)
        workbook_file.writestr(
            "xl/worksheets/sheet1.xml",
            worksheet_xml(rows, widths, len(rows[0])),
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
            create_csv(data["config"], profiles, output_path)
        else:
            create_workbook(data["config"], profiles, output_path)
    except (OSError, json.JSONDecodeError, ValueError) as error:
        parser.error(str(error))

    print(f"Created {output_path} ({len(profiles)} profiles, {args.format.upper()})")


if __name__ == "__main__":
    main()
