#pragma once

#include <cstdint>
#include <stdint.h>

#include "pico/stdlib.h"
#include "hardware/i2c.h"

namespace santroller {

/**
 * @brief Utility class to scan an I2C bus for connected slave devices on RP2040 hardware.
 */
class I2CScanner {
public:
    /**
     * @brief Checks if an I2C address is reserved per the I2C specification.
     * @param addr 7-bit I2C address (0x00 to 0x7F)
     * @return true if reserved (0x00-0x07 or 0x78-0x7F), false otherwise.
     */
    static bool is_reserved_address(uint8_t addr);

    /**
     * @brief Probes a specific 7-bit address to see if a device acknowledges.
     * @param i2c Hardware I2C instance (i2c0 or i2c1)
     * @param addr 7-bit address to probe
     * @return true if device responded with ACK, false on NACK/timeout.
     */
    static bool probe_address(i2c_inst_t* i2c, uint8_t addr);

    /**
     * @brief Scans the specified I2C instance for responding devices and prints a matrix via stdout.
     * @param i2c Hardware I2C instance (i2c0 or i2c1)
     * @param sda_pin GPIO pin mapped to SDA
     * @param scl_pin GPIO pin mapped to SCL
     * @param baudrate I2C clock speed in Hz (default 100000 Hz / 100 kHz for probing)
     * @return Number of detected devices on the bus.
     */
    static int scan_bus(i2c_inst_t* i2c,
                        uint8_t sda_pin,
                        uint8_t scl_pin,
                        uint32_t baudrate = 100000);
};

}  // namespace santroller
