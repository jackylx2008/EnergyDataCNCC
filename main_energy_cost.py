"""
能耗费用数据处理主程序
==========================

从能耗费用角度处理能源台账 Excel 文件。
1. 读取 config.yaml 中配置的输入文件。
2. 生成汇总报表 (output/energy_summary_能耗费用.xlsx)。
3. 生成统计图表 (output/charts_能耗费用/)。
"""

import logging
import os
import shutil
import time
from logging_config import setup_logger
from process_energy_data import load_config, process_energy_cost_workflow
from generate_charts import (
    generate_pie_charts,
    generate_cost_bar_chart,
    generate_grouped_bar_chart,
)


def main():
    # 设置主日志
    logger = setup_logger(
        log_level=logging.INFO, log_file="./logs/main_energy_cost.log"
    )
    logger.info("开始执行 [能耗费用] 数据处理工作流...")

    config = load_config()
    # input_dir = config["paths"]["input_dir"] # Deprecated
    input_file = config["paths"]["input_file"]  # New Input File
    output_dir = config["paths"]["output_dir"]

    # 清空输出目录
    if os.path.exists(output_dir):
        logger.info(f"正在清空输出目录: {output_dir}")
        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                logger.error(f"无法删除 {file_path}: {e}")

        logger.info("清空完成，休息 2 秒...")
        time.sleep(2)
    else:
        os.makedirs(output_dir)

    print("正在处理能耗费用数据...")

    # 1. 处理数据
    # process_energy_cost_workflow 返回 {group_name: summary_df}
    results = process_energy_cost_workflow(input_file, output_dir)

    if not results:
        logger.warning("未找到有效的数据或处理失败。")
        print("未获取到汇总数据。")
        return

    # 2. 生成图表
    for group_name, summary_df in results.items():
        print(f"正在生成图表，分组: {group_name}...")
        group_charts_dir = f"{output_dir}/charts_{group_name}"

        # 生成费用分布饼图
        generate_pie_charts(
            input_file=summary_df,
            output_dir=group_charts_dir,
            suffix="_费用(元)",
            title_prefix="能源费用分布",
        )

        # 生成标准煤(吨标准煤)分布饼图
        generate_pie_charts(
            input_file=summary_df,
            output_dir=group_charts_dir,
            suffix="_标准煤(吨标准煤)",
            title_prefix="能源标准煤(吨标准煤)分布",
        )

        generate_cost_bar_chart(input_file=summary_df, output_dir=group_charts_dir)
        generate_grouped_bar_chart(input_file=summary_df, output_dir=group_charts_dir)

        logger.info(f"分组 [{group_name}] 图表生成完成。")
        print(f"图表已保存至: {group_charts_dir}")

    print("\n" + "=" * 50)
    print("能耗费用数据处理工作流执行完成！")
    print("=" * 50)


if __name__ == "__main__":
    main()
