# Sensor-Manual-Reading-For-testing-
# STM32 CO₂ & TVOC Data Logger – Python Serial to Excel

This Python script reads real-time air quality data (CO₂ and TVOC) from a serial connection (e.g., STM32 microcontroller via USB), extracts the values, and appends them to an Excel spreadsheet with timestamps.

---

##  Features

- Extracts `CO₂ (ppm)` and `TVOC (ppb)` from serial output
- Parses and logs data with timestamps
- Stores all data in an Excel spreadsheet (`.xlsx`)
- Automatically creates directories and file if not present
- Human-readable console feedback

---

##  Requirements

- Python 3.8+
- Libraries:
  - `pandas`
  - `openpyxl`
  - `pyserial`
  - `re`
- Hardware:
  - STM32 or any microcontroller sending CO₂/TVOC over UART serial

---

##  Serial Configuration

- **Port**: `COM5` *(Update based on your STM32 COM port)*
- **Baud Rate**: `115200`
- **Timeout**: `1 second`

---

## File Output

- Excel file is saved to:  
  `C:\sensor_reading\sensor_reading.xlsx`

- Columns:
  - `Timestamp` – Date and time of each reading
  - `CO₂ (ppm)`
  - `TVOC (ppb)`

---

##  How to Run

1. Connect your STM32 board and ensure it’s outputting UART sensor data.
2. Open terminal or run the script in an IDE.
3. Script will log data live
