import pandas as pd
import os

output_file = "./output/energy_usage_summary.xlsx"
if os.path.exists(output_file):
    df = pd.read_excel(output_file)
    print(df.to_string())
else:
    print("Output file not found.")
