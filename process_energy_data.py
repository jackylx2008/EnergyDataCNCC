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
        "采暖热表": 10,  # 热力-成本
        "燃气": 8,  # 天然气-成本
        "自来水": 15,  # 自来水-成本
        "生活热水表": 17,  # 生活热水-成本
    }

    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # 寻找 '1月' 所在的行
        start_row = -1
        for i, row in df_raw.iterrows():
            if str(row[0]).strip() == "1月":
                start_row = i
                break

        if start_row == -1:
            return None

        for i in range(int(start_row), int(start_row) + 12):
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

    # Reorder columns based on specific order
    ordered_columns = ["日期区间"]
    target_order = ["电", "采暖热表", "生活热水表", "自来水", "中水", "燃气"]

    for energy_type in target_order:
        col_cost = f"{energy_type}_费用(元)"
        if col_cost in pivot_df.columns:
            ordered_columns.append(col_cost)

    remaining_cols = [col for col in pivot_df.columns if col not in ordered_columns]
    ordered_columns.extend(remaining_cols)
    pivot_df = pivot_df[ordered_columns]

    # Calculate total cost
    cost_cols = [col for col in pivot_df.columns if col.endswith("_费用(元)")]
    pivot_df["总费用(元)"] = pivot_df[cost_cols].sum(axis=1)

    output_path = os.path.join(output_dir, f"energy_summary_{group_name}.xlsx")
    pivot_df.to_excel(output_path, index=False)
    logger.info(f"分组 {group_name} 汇总已保存至 {output_path}")

    return output_path


def process_2025_workflow(input_dir, output_dir):
    """
    工作流 1: 处理 2025 年度台账文件
    """
    logger = logging.getLogger(__name__)
    group_name = "2025台账"
    summary_data = []

    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx") and f.startswith("2025") and not f.startswith("~$")
    ]

    for file_name in files:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"[2025台账流] 正在处理文件: {file_name}")

        ledger_df = process_ledger_file(file_path)
        if ledger_df is not None:
            summary_data.append(ledger_df)

    path = save_summary(summary_data, output_dir, group_name)
    return {group_name: path} if path else {}


def process_phase2_workflow(input_dir, output_dir):
    """
    工作流 2: 处理国会二期能源结算表
    """
    logger = logging.getLogger(__name__)
    group_name = "国会二期"
    summary_data = []

    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx")
        and f.startswith("国会二期主体能源结算计量表")
        and not f.startswith("~$")
    ]

    for file_name in files:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"[国会二期流] 正在处理文件: {file_name}")

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
