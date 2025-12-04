import pandas as pd
import os

input_dir = "./input"
files = [
    f for f in os.listdir(input_dir) if f.endswith(".xlsx") and not f.startswith("~$")
]

if not files:
    print("No Excel files found.")
else:
    file_path = os.path.join(input_dir, files[0])
    print(f"Inspecting file: {file_path}")

    try:
        # Read all sheet names
        xls = pd.ExcelFile(file_path)
        print(f"Sheet names: {xls.sheet_names}")

        # Read first sheet
        df = pd.read_excel(file_path, sheet_name=0)

        # Forward fill '能源类型' to handle merged cells
        df["能源类型"] = df["能源类型"].ffill()

        print("\nFirst Sheet Columns:")
        print(df.columns.tolist())
        print("\nAll rows:")
        print(df.to_string())

    except Exception as e:
        print(f"Error reading excel: {e}")
