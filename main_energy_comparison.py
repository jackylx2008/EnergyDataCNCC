"""
年度能耗对比工作流
==========================

用于比较不同年份（例如 2024 vs 2025）相同月份的能耗费用和标准煤耗量。
1. 读取 config.yaml 中 year_comparison 的配置。
2. 调用 process_energy_data 中的逻辑处理多个年份的 Excel 台账。
3. 调用 generate_charts 中的逻辑生成对比统计图表。

使用说明:
    请确保 config.yaml 中已配置 year_comparison.files，例如:
    year_comparison:
      files:
        2024: path/to/2024_data.xlsx
        2025: path/to/2025_data.xlsx
"""

import logging
import os
import shutil
from core.logging_config import setup_logger
from core.process_energy_data import load_config, process_multi_year_comparison
from core.generate_charts import (
    generate_multi_energy_comparison_bar,
    generate_comparison_pie_charts,
)


def main():
    # 设置对比工作流专用日志
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/main_comparison.log")
    logger.info("开始执行 [年度对比] 数据处理工作流...")

    # 加载配置
    config = load_config()
    comparison_config = config.get("year_comparison", {})
    year_file_map = comparison_config.get("files", {})
    output_dir = config["paths"]["output_dir"]
    comparison_output_dir = os.path.join(output_dir, "charts_年度对比")

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

    if not year_file_map:
        logger.error("config.yaml 中未找到 year_comparison.files 配置项。")
        print("Error: 请在 config.yaml 中配置 year_comparison.files 路径。")
        return

    # 1. 数据处理：汇总之不同年份的数据，并计算标准煤
    print("正在处理多年份数据...")
    df_combined = process_multi_year_comparison(year_file_map, output_dir)

    if df_combined is None or df_combined.empty:
        logger.warning("处理后未获取到有效对比数据。")
        print("未获取到有效的对比数据，请检查输入文件路径和格式。")
        return

    # 根据 config.yaml 中的 months 进行过滤
    comparison_months = comparison_config.get("months", [])
    if comparison_months:
        # 将 "1" 转换为 "1月" 的格式进行匹配
        target_months = [f"{m}月" for m in comparison_months]
        df_combined = df_combined[df_combined["日期区间"].isin(target_months)]
        logger.info(f"已根据配置过滤月份: {target_months}")

        if df_combined.empty:
            logger.warning(f"过滤月份 {target_months} 后未获取到有效对比数据。")
            print(f"警告: 在指定的月份 {target_months} 内未找到数据。")
            return

    # 2. 图表生成
    print(f"正在生成同比分析图表，保存至: {comparison_output_dir}...")

    # A. 多能源类型在相同月份的对比 (用户需求: 在一个PNG中比较不同能源)
    print("正在生成多能源汇总对比图...")
    generate_multi_energy_comparison_bar(
        df=df_combined,
        output_dir=comparison_output_dir,
        value_col="费用(元)",
        title="年度能源费用同比细项对比",
        ylabel="费用",
        filename="comparison_all_energies_cost.png",
        use_wan_yuan=True,
    )
    generate_multi_energy_comparison_bar(
        df=df_combined,
        output_dir=comparison_output_dir,
        value_col="标准煤(吨标准煤)",
        title="年度能源标准煤同比细项对比",
        ylabel="标准煤 (吨)",
        filename="comparison_all_energies_coal.png",
        use_wan_yuan=False,
    )

    # C. 同月份不同年份的饼图对比 (用户需求)
    print("正在生成同比构成饼图...")
    generate_comparison_pie_charts(
        df=df_combined,
        output_dir=comparison_output_dir,
        value_col="费用(元)",
        title_prefix="能源费用构成同比",
        filename="comparison_pie_cost_ratio.png",
    )
    generate_comparison_pie_charts(
        df=df_combined,
        output_dir=comparison_output_dir,
        value_col="标准煤(吨标准煤)",
        title_prefix="能源标准煤构成同比",
        filename="comparison_pie_coal_ratio.png",
    )

    # B. 记录日志完成
    logger.info("年度对比工作流执行完成。")
    print("\n--- 年度对比处理完成 ---")
    print(f"对比数据明细: {os.path.join(output_dir, 'energy_comparison_details.xlsx')}")
    print(f"对比分析图表: {comparison_output_dir}")


if __name__ == "__main__":
    main()
