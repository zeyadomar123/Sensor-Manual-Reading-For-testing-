##final format of saving to excel
import serial
import datetime
import pandas as pd
import os
import re

# Define the file path
xlsx_filename = "C:\\sensor_reading\\sensor_reading.xlsx"

# Ensure the directory exists
os.makedirs(os.path.dirname(xlsx_filename), exist_ok=True)

# Configure Serial Port (Change 'COM5' to match your STM32's COM port)
ser = serial.Serial('COM5', 115200, timeout=1)

print(f"Saving sensor data to: {xlsx_filename}")
print("Reading sensor data... Press Ctrl+C to stop.")

def extract_values(sensor_string):
    """ Extracts CO₂ and TVOC values from the sensor reading string. """
    co2_match = re.search(r"CO2:\s*(\d+)\s*ppm", sensor_string)  # Fix CO₂ regex
    tvoc_match = re.search(r"TVOC:\s*(\d+)\s*ppb", sensor_string)

    co2 = int(co2_match.group(1)) if co2_match else 0  # Default to 0 instead of None
    tvoc = int(tvoc_match.group(1)) if tvoc_match else 0

    return co2, tvoc

try:
    while True:
        line = ser.readline().decode('utf-8', errors='ignore').strip()
        if line:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            co2, tvoc = extract_values(line)

            print(f"[{timestamp}] CO₂: {co2} ppm, TVOC: {tvoc} ppb")

            # Load existing data if the file exists, otherwise create an empty DataFrame
            if os.path.exists(xlsx_filename):
                df = pd.read_excel(xlsx_filename, engine="openpyxl")
            else:
                df = pd.DataFrame(columns=["Timestamp", "CO₂ (ppm)", "TVOC (ppb)"])

            # Append new data
            new_data = pd.DataFrame([[timestamp, co2, tvoc]], columns=["Timestamp", "CO₂ (ppm)", "TVOC (ppb)"])
            df = pd.concat([df, new_data], ignore_index=True)

            # Save updated data to Excel
            df.to_excel(xlsx_filename, index=False, engine="openpyxl")

            print("✅ Data successfully written to Excel.")
except KeyboardInterrupt:
    print("\nExiting... Data saved to", xlsx_filename)
    ser.close()
