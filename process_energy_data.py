"""
能源数据处理主程序
================

此脚本负责协调整个能源数据的处理流程。主要功能包括：
1. 读取配置文件 (config.yaml)。
2. 扫描输入目录中的 Excel 文件。
3. 使用 EnergySheet 类处理每个工作表的数据。
4. 管理数据缓存 (Parquet 格式)，避免重复处理并检查数据一致性。
5. 生成汇总 Excel 报表，包含各能源类型的费用统计。

使用方法:
    直接运行此脚本: python process_energy_data.py
"""

import os
import yaml
import pandas as pd
import logging
from logging_config import setup_logger
from energy_models import EnergySheet


def process_ledger_file(file_path):
    """
    处理 2025 台账格式的 Excel 文件。
    返回一个包含汇总数据的 list of DataFrames。
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)

    mapping = {
        "电": 2,
        "燃气": 8,  # 天然气-成本
        "自来水": 15,  # 自来水-成本
    }

    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # 寻找 '1月' 所在的行
        start_row = -1
        for i in range(len(df_raw)):
            if str(df_raw.iloc[i, 0]).strip() == "1月":
                start_row = i
                break

        if start_row == -1:
            return None

        for i in range(start_row, start_row + 12):
            if i >= len(df_raw):
                break
            row = df_raw.iloc[i]
            month = str(row[0]).strip()
            if not month or "月" not in month or len(month) > 3:
                continue

            for energy_type, col_idx in mapping.items():
                val = 0.0
                if col_idx < len(row):
                    try:
                        raw_val = row[col_idx]
                        val = float(raw_val)
                        if pd.isna(val) or val < 0:
                            val = 0.0
                    except (ValueError, TypeError):
                        val = 0.0

                # Fix for '自来水' (Tap Water): The Excel file often uses a formula (=O*9.5)
                # which isn't cached, resulting in NaN or 0.
                if energy_type == "自来水" and val == 0:
                    try:
                        # Index 14 is Tap Water Volume
                        vol_val = float(row[14])
                        if not pd.isna(vol_val) and vol_val > 0:
                            val = vol_val * 9.5
                    except (ValueError, TypeError, IndexError):
                        pass

                # Always append, even if 0, to ensure columns appear in summary
                results.append(
                    {
                        "日期区间": month,
                        "能源类型": energy_type,
                        "实际消耗": 0.0,
                        "费用(元)": val,
                        "来源文件": file_name,
                    }
                )
        return pd.DataFrame(results)
    except Exception as e:
        logger.error(f"处理台账文件 {file_name} 失败: {e}")
        return None


def load_config(config_path="config.yaml"):
    """
    加载 YAML 配置文件。

    Args:
        config_path (str): 配置文件路径，默认为 "config.yaml"。

    Returns:
        dict: 包含配置信息的字典。
    """
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def save_summary(summary_data, output_dir, group_name):
    """
    Helper function to save summary data for a group.
    """
    logger = logging.getLogger(__name__)

    if not summary_data:
        return None

    logger.info(f"正在汇总分组数据: {group_name}")
    final_df = pd.concat(summary_data, ignore_index=True)

    # Pivot the table to have Energy Types as headers
    pivot_df = final_df.pivot_table(
        index=["日期区间"],
        columns="能源类型",
        values=["费用(元)"],
        aggfunc="sum",
        fill_value=0,
    )

    # Swap levels to group by Energy Type (e.g., 电_费用)
    pivot_df.columns = pivot_df.columns.swaplevel(0, 1)
    pivot_df.sort_index(axis=1, level=0, inplace=True)

    # Flatten MultiIndex columns
    pivot_df.columns = [f"{col[0]}_{col[1]}" for col in pivot_df.columns]

    # Reset index to make '日期区间' a normal column
    pivot_df.reset_index(inplace=True)

    # 排序逻辑: 尝试按月份数字排序
    def sort_key(x):
        if isinstance(x, str):
            import re

            match = re.search(r"(\d+)", x)
            if match:
                return int(match.group(1))
        return 999

    pivot_df["_sort_key"] = pivot_df["日期区间"].apply(sort_key)
    pivot_df.sort_values("_sort_key", inplace=True)
    pivot_df.drop(columns=["_sort_key"], inplace=True)

    # Reorder columns based on specific order, and exclude those with 0 total cost
    ordered_columns = ["日期区间"]
    target_order = ["电", "采暖热费", "生活热水热费", "自来水", "中水", "燃气"]

    for energy_type in target_order:
        col_cost = f"{energy_type}_费用(元)"
        if col_cost in pivot_df.columns:
            # Check if this category has any non-zero costs
            if pivot_df[col_cost].sum() > 0:
                ordered_columns.append(col_cost)

    remaining_cols = [
        col
        for col in pivot_df.columns
        if col not in ordered_columns and pivot_df[col].sum() > 0
    ]
    ordered_columns.extend(remaining_cols)
    pivot_df = pivot_df[ordered_columns]

    # Calculate total cost
    cost_cols = [col for col in pivot_df.columns if col.endswith("_费用(元)")]
    pivot_df["总费用(元)"] = pivot_df[cost_cols].sum(axis=1)

    output_path = os.path.join(output_dir, f"energy_summary_{group_name}.xlsx")
    pivot_df.to_excel(output_path, index=False)
    logger.info(f"分组 {group_name} 汇总已保存至 {output_path}")

    return output_path


def process_reclaimed_water_file(file_path):
    """
    处理中水用量表。
    1. 月份在第0列 (如 1月, 2月)。
    2. 费用在第2列。
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)
    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        for i, row in df_raw.iterrows():
            month = str(row[0]).strip()
            if "月" not in month or len(month) > 3:
                continue

            try:
                # B列(索引1)为用量，C列(索引2)为费用
                vol = 0.0
                cost = 0.0

                def convert_to_float(val):
                    if pd.isna(val):
                        return None
                    try:
                        return float(val)
                    except:
                        return None

                raw_vol = convert_to_float(row[1]) if len(row) > 1 else None
                raw_cost = convert_to_float(row[2]) if len(row) > 2 else None

                if raw_vol is not None:
                    vol = raw_vol

                if raw_cost is not None:
                    cost = raw_cost
                else:
                    # 如果费用列为空，按 1.0 元/立方米估算
                    cost = vol * 1.0

                if vol >= 0:
                    results.append(
                        {
                            "日期区间": month,
                            "能源类型": "中水",
                            "实际消耗": vol,
                            "费用(元)": cost,
                            "来源文件": file_name,
                        }
                    )
            except (ValueError, TypeError):
                pass

        return pd.DataFrame(results)
    except Exception as e:
        logger.error(f"处理中水文件 {file_name} 失败: {e}")
        return None


def process_heat_station_file(file_path):
    """
    处理热力站月度费用。
    1. 月份在第0列 (1月, 2月, ..., 12月)。
    2. 生活热水热费在第1列。
    3. 采暖热费在第2列。
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)
    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=0)

        for i, row in df_raw.iterrows():
            month = str(row.iloc[0]).strip()
            if "月" not in month or "合计" in month or month == "月份":
                continue

            # 生活热水热费 (Col 1)
            try:
                hw_val = float(row.iloc[1])
                if not pd.isna(hw_val) and hw_val >= 0:
                    results.append(
                        {
                            "日期区间": month,
                            "能源类型": "生活热水热费",
                            "实际消耗": 0.0,
                            "费用(元)": hw_val,
                            "来源文件": file_name,
                        }
                    )
            except (ValueError, TypeError):
                pass

            # 采暖热费 (Col 2)
            try:
                h_val = float(row.iloc[2])
                if not pd.isna(h_val) and h_val >= 0:
                    results.append(
                        {
                            "日期区间": month,
                            "能源类型": "采暖热费",
                            "实际消耗": 0.0,
                            "费用(元)": h_val,
                            "来源文件": file_name,
                        }
                    )
            except (ValueError, TypeError):
                pass

        return pd.DataFrame(results)
    except Exception as e:
        logger.error(f"处理热力站文件 {file_name} 失败: {e}")
        return None


def process_2025_workflow(input_dir, output_dir):
    """
    工作流 1: 处理 2025 年度台账及中水文件
    """
    logger = logging.getLogger(__name__)
    group_name = "2025台账"
    summary_data = []

    # 处理主要的 2025 台账文件
    files_ledger = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx")
        and ("台账" in f or f.startswith("2025"))
        and not f.startswith("~$")
        and "热力站" not in f  # 排除专门的热力站费用文件
    ]

    for file_name in files_ledger:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"[2025台账流] 正在处理台账文件: {file_name}")

        ledger_df = process_ledger_file(file_path)
        if ledger_df is not None:
            summary_data.append(ledger_df)

    # 包含 B23 目录下的台账文件 (如果是 2025 台账)
    base_dir = (
        os.path.dirname(input_dir)
        if os.path.basename(input_dir) == "主体"
        else input_dir
    )
    b23_dir = os.path.join(base_dir, "B23")
    if os.path.exists(b23_dir):
        files_b23 = [
            f
            for f in os.listdir(b23_dir)
            if f.endswith(".xlsx")
            and ("台账" in f or f.startswith("2025"))
            and not f.startswith("~$")
        ]
        for file_name in files_b23:
            file_path = os.path.join(b23_dir, file_name)
            logger.info(f"[2025台账流] 正在处理 B23 台账文件: {file_name}")
            ledger_df = process_ledger_file(file_path)
            if ledger_df is not None:
                summary_data.append(ledger_df)

    # 处理特定的热力站费用文件 (递归搜索)
    files_heat = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if (
                "热力站供暖费与生活热水费" in f
                and f.endswith(".xlsx")
                and not f.startswith("~$")
            ):
                files_heat.append(os.path.join(root, f))

    for file_path in files_heat:
        file_name = os.path.basename(file_path)
        logger.info(f"[2025台账流] 正在处理热力文件: {file_name}")
        heat_df = process_heat_station_file(file_path)
        if heat_df is not None:
            summary_data.append(heat_df)

    # 处理中水文件 (递归搜索)
    files_reclaimed = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if "中水用量表" in f and f.endswith(".xlsx") and not f.startswith("~$"):
                files_reclaimed.append(os.path.join(root, f))

    for file_path in files_reclaimed:
        file_name = os.path.basename(file_path)
        logger.info(f"[2025台账流] 正在处理中水文件: {file_name}")
        reclaimed_df = process_reclaimed_water_file(file_path)
        if reclaimed_df is not None:
            summary_data.append(reclaimed_df)

    path = save_summary(summary_data, output_dir, group_name)
    return {group_name: path} if path else {}


def process_phase2_workflow(input_dir, output_dir):
    """
    工作流 2: 处理国会二期能源结算表
    """
    logger = logging.getLogger(__name__)
    group_name = "国会二期"
    summary_data = []

    # 1. 处理主体结算表 (国会二期主体)
    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx")
        and f.startswith("国会二期主体能源结算计量表")
        and not f.startswith("~$")
    ]

    for file_name in files:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"[国会二期流] 正在处理主体文件: {file_name}")

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                logger.info(f"  正在处理工作表: {sheet_name}")
                sheet_obj = EnergySheet(file_path, sheet_name)

                # Check cache
                cache_dir = os.path.join(os.path.dirname(output_dir), "data")
                comparison_result = sheet_obj.compare_with_cache(
                    cache_dir, format="parquet"
                )

                if comparison_result == "NEW":
                    logger.info(f"检测到新工作表: {sheet_name}。正在保存到缓存。")
                    sheet_obj.save_data(cache_dir, format="parquet")
                elif comparison_result == "MATCH":
                    logger.info(f"工作表 {sheet_name} 数据一致性检查通过。")  # OK

                summary = sheet_obj.get_summary()
                if summary is not None:
                    summary_data.append(summary)
        except Exception as e:
            logger.error(f"处理文件 {file_name} 失败: {e}", exc_info=True)

    # 2. 处理国会二期相关的中水文件
    base_dir = (
        os.path.dirname(input_dir)
        if os.path.basename(input_dir) == "主体"
        else input_dir
    )
    files_reclaimed = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            # 仅处理文件名包含 "国会二期" 和 "中水用量表" 的文件
            if (
                "中水用量表" in f
                and "国会二期" in f
                and f.endswith(".xlsx")
                and not f.startswith("~$")
            ):
                files_reclaimed.append(os.path.join(root, f))

    for file_path in files_reclaimed:
        file_name = os.path.basename(file_path)
        logger.info(f"[国会二期流] 正在处理中水文件: {file_name}")
        reclaimed_df = process_reclaimed_water_file(file_path)
        if reclaimed_df is not None:
            summary_data.append(reclaimed_df)

    path = save_summary(summary_data, output_dir, group_name)
    return {group_name: path} if path else {}


def process_excel_files():
    """
    主入口: 依次执行各类独立工作流
    """
    config = load_config()

    # Setup logging
    log_level_str = config.get("log_level", "INFO")
    log_level = getattr(logging, log_level_str.upper(), logging.INFO)
    log_file = config.get("log_file", "./logs/app.log")

    # 这里 setup_logger 只需调用一次，内部实现如果是单例模式则没问题
    # 如果会重复添加 handler，可能需要注意。这里假设 setup_logger 是安全的。
    logger = setup_logger(log_level=log_level, log_file=log_file)

    input_dir = config["paths"]["input_dir"]
    output_dir = config["paths"]["output_dir"]

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    all_summaries = {}

    # 执行 2025 台账工作流
    try:
        res1 = process_2025_workflow(input_dir, output_dir)
        all_summaries.update(res1)
    except Exception as e:
        logger.error(f"2025 台账工作流执行失败: {e}")

    # 执行 国会二期工作流
    try:
        res2 = process_phase2_workflow(input_dir, output_dir)
        all_summaries.update(res2)
    except Exception as e:
        logger.error(f"国会二期工作流执行失败: {e}")

    return all_summaries


if __name__ == "__main__":
    process_excel_files()
