#include "i2c_scanner.hpp"

#include <cstdio>
#include <stdint.h>

namespace santroller {

bool I2CScanner::is_reserved_address(uint8_t addr) {
    // Addresses 0x00..0x07 and 0x78..0x7F are reserved by the I2C specification.
    return (addr & 0x78) == 0 || (addr & 0x78) == 0x78;
}

bool I2CScanner::probe_address(i2c_inst_t* i2c, uint8_t addr) {
    // Perform a 1-byte dummy read to check for ACK from target device.
    uint8_t rxdata = 0;
    int result = i2c_read_blocking(i2c, addr, &rxdata, 1, false);
    return result >= 0;
}

int I2CScanner::scan_bus(i2c_inst_t* i2c, uint8_t sda_pin, uint8_t scl_pin, uint32_t baudrate) {
    // Initialize hardware I2C peripheral and configure GPIO functions.
    i2c_init(i2c, baudrate);
    gpio_set_function(sda_pin, GPIO_FUNC_I2C);
    gpio_set_function(scl_pin, GPIO_FUNC_I2C);
    gpio_pull_up(sda_pin);
    gpio_pull_up(scl_pin);

    printf("\n--- Starting Santroller I2C Bus Scan ---\n");
    printf("Bus: %s | SDA: GP%d | SCL: GP%d | Frequency: %lu Hz\n\n",
           (i2c == i2c0) ? "i2c0" : "i2c1",
           sda_pin,
           scl_pin,
           static_cast<unsigned long>(baudrate));

    printf("     0  1  2  3  4  5  6  7  8  9  A  B  C  D  E  F\n");

    int devices_found = 0;

    for (int addr = 0; addr < 128; addr += 16) {
        printf("%02X: ", addr);
        for (int col = 0; col < 16; ++col) {
            uint8_t current_addr = static_cast<uint8_t>(addr + col);

            if (is_reserved_address(current_addr)) {
                printf(" R ");
            } else if (probe_address(i2c, current_addr)) {
                printf("%02X ", current_addr);
                devices_found++;
            } else {
                printf(" . ");
            }
        }
        printf("\n");
    }

    printf("\nScan complete. Devices found: %d\n", devices_found);
    printf("-----------------------------------------\n\n");

    return devices_found;
}

}  // namespace santroller
