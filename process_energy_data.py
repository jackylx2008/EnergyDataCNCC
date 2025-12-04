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


def process_excel_files():
    """
    主处理函数。

    执行以下步骤:
    1. 初始化日志和配置。
    2. 遍历输入目录下的所有 Excel 文件。
    3. 对每个工作表进行清洗、缓存比对和汇总。
    4. 将所有汇总数据合并，并生成透视表。
    5. 将最终结果保存为 Excel 文件。
    """
    # Load configuration
    config = load_config()

    # Setup logging
    log_level_str = config.get("log_level", "INFO")
    log_level = getattr(logging, log_level_str.upper(), logging.INFO)
    log_file = config.get("log_file", "./logs/app.log")
    logger = setup_logger(log_level=log_level, log_file=log_file)

    input_dir = config["paths"]["input_dir"]
    output_dir = config["paths"]["output_dir"]

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"已创建输出目录: {output_dir}")

    summary_data = []

    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not files:
        logger.warning("输入目录中未找到 Excel 文件。")
        return

    for file_name in files:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"正在处理文件: {file_name}")

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                logger.info(f"  正在处理工作表: {sheet_name}")

                # Use the EnergySheet class
                sheet_obj = EnergySheet(file_path, sheet_name)

                # Compare with cache and save if new
                cache_dir = os.path.join(os.path.dirname(output_dir), "data")
                comparison_result = sheet_obj.compare_with_cache(
                    cache_dir, format="parquet"
                )

                if comparison_result == "NEW":
                    logger.info(f"检测到新工作表: {sheet_name}。正在保存到缓存。")
                    sheet_obj.save_data(cache_dir, format="parquet")
                elif comparison_result == "MATCH":
                    logger.info(f"工作表 {sheet_name} 数据一致性检查通过。")
                elif comparison_result == "MISMATCH":
                    logger.error(
                        f"文件 {file_name} 中的工作表 {sheet_name} "
                        "与缓存数据不匹配！跳过缓存更新。"
                    )
                else:
                    logger.error(f"检查工作表 {sheet_name} 的缓存时出错。")

                summary = sheet_obj.get_summary()

                if summary is not None:
                    summary_data.append(summary)

        except Exception as e:
            logger.error(f"处理文件 {file_name} 失败: {e}", exc_info=True)

    if summary_data:
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

        # Reorder columns based on specific order
        ordered_columns = ["日期区间"]
        target_order = ["电", "采暖热表", "生活热水表", "自来水", "中水", "燃气"]

        for energy_type in target_order:
            col_cost = f"{energy_type}_费用(元)"
            if col_cost in pivot_df.columns:
                ordered_columns.append(col_cost)

        # Add any remaining columns that might not be in the target list (just in case)
        remaining_cols = [col for col in pivot_df.columns if col not in ordered_columns]
        ordered_columns.extend(remaining_cols)

        pivot_df = pivot_df[ordered_columns]

        # Calculate total cost
        cost_cols = [col for col in pivot_df.columns if col.endswith("_费用(元)")]
        pivot_df["总费用(元)"] = pivot_df[cost_cols].sum(axis=1)

        output_path = os.path.join(output_dir, "energy_usage_summary.xlsx")
        pivot_df.to_excel(output_path, index=False)
        logger.info(f"汇总已保存至 {output_path}")
        print(f"处理完成。汇总已保存至 {output_path}")
    else:
        logger.warning("未处理任何数据。")
        print("未处理任何数据。")


if __name__ == "__main__":
    process_excel_files()
