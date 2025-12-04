"""
Excel 文件检查工具
================

这是一个辅助脚本，用于快速查看输入目录中 Excel 文件的结构和内容。
主要用于开发调试阶段，确认 Excel 文件的列名、Sheet 名称以及数据格式。

功能:
1. 列出 input 目录下的 Excel 文件。
2. 读取第一个文件的所有 Sheet 名称。
3. 读取第一个 Sheet 的前几行数据并打印，用于检查数据清洗前的原始状态。
"""

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
