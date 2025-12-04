"""
主程序入口
==========

整合能源数据处理与图表生成的工作流。
按顺序执行以下步骤：
1. 处理 Excel 数据并生成汇总报表 (process_energy_data.py)。
2. 基于汇总报表生成各类统计图表 (generate_charts.py)。

使用方法:
    python main.py
"""

import logging
from logging_config import setup_logger
from process_energy_data import process_excel_files
from generate_charts import (
    generate_pie_charts,
    generate_cost_bar_chart,
    generate_grouped_bar_chart,
)


def main():
    """
    执行完整的工作流：
    1. 数据清洗与汇总
    2. 可视化图表生成
    """
    # 设置主日志
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/main.log")
    logger.info("开始执行能源数据处理工作流...")

    try:
        # 1. 处理数据
        logger.info(">>> 阶段 1: 处理 Excel 数据...")
        print("正在处理 Excel 数据...")
        process_excel_files()
        logger.info("数据处理完成。")

        # 2. 生成图表
        logger.info(">>> 阶段 2: 生成统计图表...")
        print("正在生成统计图表...")
        generate_pie_charts()
        generate_cost_bar_chart()
        generate_grouped_bar_chart()
        logger.info("图表生成完成。")

        logger.info("工作流执行成功！所有结果已保存至 output 目录。")
        print("\n" + "=" * 50)
        print("工作流执行成功！所有结果已保存至 output 目录。")
        print("=" * 50)

    except Exception as e:
        logger.error(f"工作流执行失败: {e}", exc_info=True)
        print(f"\n工作流执行失败: {e}")


if __name__ == "__main__":
    main()
