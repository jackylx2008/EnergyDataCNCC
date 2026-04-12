import os
from pathlib import Path

import pandas as pd
import numpy as np


def get_month_index(month_str):
    if not isinstance(month_str, str):
        return None
    import re

    match = re.search(r"(\d+)", month_str)
    if match:
        return int(match.group(1))
    return None


def process_data():
    project_root = Path(__file__).resolve().parent.parent
    input_dir = str(project_root / "input")
    output_dir = str(project_root / "output")

    # Initialize storage for 1-12 months
    # Categories: 电, 采暖热, 生活热水, 自来水, 中水, 燃气
    months = [f"{i}月" for i in range(1, 13)]
    cats = ["电", "采暖热", "生活热水", "自来水", "中水", "燃气"]
    data = {
        m: {f"{cat}_{type_}": 0.0 for cat in cats for type_ in ["用量", "费用"]}
        for m in months
    }

    # 1. Process Ledger Files (Exclude B23)
    ledger_files = [
        f
        for f in os.listdir(input_dir)
        if "台账" in f
        and f.endswith(".xlsx")
        and not f.startswith("~$")
        and "B23" not in f  # Exclude B23
    ]

    for f in ledger_files:
        path = os.path.join(input_dir, f)
        print(f"Processing ledger: {f}")
        df_raw = pd.read_excel(path, header=None)

        # Find start row
        start_row = -1
        for i in range(len(df_raw)):
            if str(df_raw.iloc[i, 0]).strip() == "1月":
                start_row = i
                break

        if start_row == -1:
            continue

        for i in range(start_row, start_row + 12):
            if i >= len(df_raw):
                break
            row = df_raw.iloc[i]
            month = str(row[0]).strip()
            if "月" not in month:
                continue

            # Map month to our data dict
            m_idx = get_month_index(month)
            if not m_idx:
                continue
            m_key = f"{m_idx}月"
            if m_key not in data:
                continue

            # Helper to safely get float
            def get_val_ledger(idx):
                try:
                    if idx >= len(row):
                        return 0.0
                    v = float(row[idx])
                    return v if not np.isnan(v) else 0.0
                except Exception:
                    return 0.0

            data[m_key]["电_用量"] += get_val_ledger(1)
            data[m_key]["电_费用"] += get_val_ledger(2)
            data[m_key]["燃气_用量"] += get_val_ledger(7)
            data[m_key]["燃气_费用"] += get_val_ledger(8)

            # For Heat: Ledger Col 9/10 is usually the Total Heat.
            # We will use the Heat Station file to split the COST later.
            # But the volume in Col 9 is generally the heating volume (GJ).
            data[m_key]["采暖热_用量"] += get_val_ledger(9)
            # We don't add cost to 采暖热_费用 yet, or we add total and correct later.
            # Let's just store the volume for now.
            # DHW volume is Col 16.
            data[m_key]["生活热水_用量"] += get_val_ledger(16)

            # Tap Water logic with fix
            vol_w = get_val_ledger(14)
            cost_w = get_val_ledger(15)
            if cost_w == 0 and vol_w > 0:
                cost_w = vol_w * 9.5

            data[m_key]["自来水_用量"] += vol_w
            data[m_key]["自来水_费用"] += cost_w

    # 2. Process Heat Station File (for split costs)
    heat_station_files = [
        f for f in os.listdir(input_dir) if "热力站" in f and f.endswith(".xlsx")
    ]
    for f in heat_station_files:
        path = os.path.join(input_dir, f)
        print(f"Processing heat station: {f}")
        df_heat = pd.read_excel(path)
        # Assuming the structure checked before: Months in Col 0, DHW Cost in Col 1, Heating Cost in Col 2
        # Data starts from row 1 (0-indexed) if header=0
        for i, row in df_heat.iterrows():
            m_str = str(row.iloc[0]).strip()
            if "月" not in m_str or "合计" in m_str:
                continue
            m_idx = get_month_index(m_str)
            if not m_idx:
                continue
            m_key = f"{m_idx}月"
            if m_key not in data:
                continue

            def get_val_row(val):
                try:
                    v = float(val)
                    return v if not np.isnan(v) else 0.0
                except Exception:
                    return 0.0

            data[m_key]["生活热水_费用"] += get_val_row(row.iloc[1])
            data[m_key]["采暖热_费用"] += get_val_row(row.iloc[2])

    # 3. Process Reclaimed Water File
    # Search recursively to include ./input/B23/
    base_dir = (
        os.path.dirname(input_dir)
        if os.path.basename(input_dir) == "主体"
        else input_dir
    )
    reclaimed_files = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if "中水用量表" in f and f.endswith(".xlsx") and not f.startswith("~$"):
                reclaimed_files.append(os.path.join(root, f))

    for path in reclaimed_files:
        df_raw = pd.read_excel(path, header=None)
        for i in range(len(df_raw)):
            row = df_raw.iloc[i]
            m_str = str(row[0]).strip()
            if "月" not in m_str or "合计" in m_str:
                continue
            m_idx = get_month_index(m_str)
            if not m_idx:
                continue
            m_key = f"{m_idx}月"

            def get_val_reclaimed(idx):
                try:
                    val = row[idx]
                    if pd.isna(val):
                        return None
                    return float(val)
                except Exception:
                    return None

            vol = get_val_reclaimed(1) or 0.0
            cost = get_val_reclaimed(2)
            if cost is None:
                # 如果费用列为空，按 1.0 元/立方米估算
                cost = vol * 1.0

            data[m_key]["中水_用量"] += vol
            data[m_key]["中水_费用"] += cost

    # Convert to DataFrame
    df_result = pd.DataFrame.from_dict(data, orient="index")
    df_result.index.name = "月份"
    df_result.reset_index(inplace=True)

    # Add Total Cost Column
    cost_cols = [col for col in df_result.columns if "费用" in col]
    df_result["总计_费用"] = df_result[cost_cols].sum(axis=1)

    # Reorder columns: 电, 采暖热, 生活热水, 自来水, 中水, 燃气, 总计
    col_order = ["月份"]
    for cat in cats:
        col_order.append(f"{cat}_用量")
        col_order.append(f"{cat}_费用")
    col_order.append("总计_费用")

    # Ensure all columns exist
    df_result = df_result[col_order]

    # Add Total Row
    total_row = df_result.sum(numeric_only=True)
    total_row["月份"] = "总计"
    # For usage columns, we sum them. For cost columns, we sum them.
    # We need to turn total_row into a list/dict that matches the df columns
    total_row_df = pd.DataFrame([total_row])
    df_final = pd.concat([df_result, total_row_df], ignore_index=True)

    # Rename columns for better appearance
    new_cols = []
    for col in df_final.columns:
        if col == "月份":
            new_cols.append("月份")
        elif col == "总计_费用":
            new_cols.append("总计 费用")
        else:
            cat, type_ = col.split("_")
            new_cols.append(f"{cat} {type_}")

    df_final.columns = new_cols

    output_path = os.path.join(output_dir, "能源数据汇总一览表_2025.xlsx")

    # Use ExcelWriter for styling if needed, but simple to_excel is fine
    df_final.to_excel(output_path, index=False)
    print(f"Consolidated table saved to: {output_path}")


if __name__ == "__main__":
    process_data()
