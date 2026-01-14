"""
2025 年度台账数据处理主程序
==========================

仅处理 2025 年度能源台账 Excel 文件。
1. 读取 input 目录下的 2025 开头的台账文件。
2. 生成汇总报表 (output/energy_summary_2025台账.xlsx)。
3. 生成统计图表 (output/charts_2025台账/)。
"""

import logging
from logging_config import setup_logger
from process_energy_data import load_config, process_2025_workflow
from generate_charts import (
    generate_pie_charts,
    generate_cost_bar_chart,
    generate_grouped_bar_chart,
)


def main():
    # 设置主日志
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/main_2025.log")
    logger.info("开始执行 [2025台账] 数据处理工作流...")

    config = load_config()
    input_dir = config["paths"]["input_dir"]
    output_dir = config["paths"]["output_dir"]

    print("正在处理 2025 台账数据...")

    # 1. 处理数据
    # process_2025_workflow 返回 {group_name: summary_path}
    results = process_2025_workflow(input_dir, output_dir)

    if not results:
        logger.warning("未找到有效的数据或处理失败。")
        print("未生成汇总数据。")
        return

    # 2. 生成图表
    for group_name, summary_file in results.items():
        print(f"正在生成图表，分组: {group_name}...")
        group_charts_dir = f"{output_dir}/charts_{group_name}"

        generate_pie_charts(input_file=summary_file, output_dir=group_charts_dir)
        generate_cost_bar_chart(input_file=summary_file, output_dir=group_charts_dir)
        generate_grouped_bar_chart(input_file=summary_file, output_dir=group_charts_dir)

        logger.info(f"分组 [{group_name}] 图表生成完成。")
        print(f"图表已保存至: {group_charts_dir}")

    print("\n" + "=" * 50)
    print("2025 台账工作流执行完成！")
    print("=" * 50)


if __name__ == "__main__":
    main()
