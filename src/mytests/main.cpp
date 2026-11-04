#include "pico/stdlib.h" 
#include "i2c_scanner.hpp" 
#include <stdio.h>

int main() 
{ 
    stdio_init_all(); 
    
    // Wait brief moment for USB CDC serial connection to open
    
    while (!stdio_usb_connected()) {
        tight_loop_contents();
    }    
    
    sleep_ms(2000); 

    // Scan i2c0 on GP0 (SDA) and GP1 (SCL) 
    int devices = santroller::I2CScanner::scan_bus(i2c0, 0, 1); 
    
    if (devices == 0) { 
        printf("Warning: No I2C devices detected! Check wiring and power.\n");
    } 
    
    while (true) { 
        tight_loop_contents(); 
    } 
}
