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
from core.logging_config import setup_logger
from core.process_energy_data import load_config, process_energy_cost_workflow
from core.generate_charts import (
    generate_pie_charts,
    generate_cost_bar_chart,
    generate_coal_bar_chart,
    generate_grouped_bar_chart,
    generate_energy_type_distribution_bar,
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
        generate_coal_bar_chart(input_file=summary_df, output_dir=group_charts_dir)
        generate_grouped_bar_chart(input_file=summary_df, output_dir=group_charts_dir)

        # 生成能源类型费用对比柱状图 (X轴为能源类型)
        generate_energy_type_distribution_bar(
            input_file=summary_df,
            output_dir=group_charts_dir,
            suffix="_费用(元)",
            title_prefix="各能源类型费用对比",
        )
        # 生成能源类型标准煤对比柱状图 (X轴为能源类型)
        generate_energy_type_distribution_bar(
            input_file=summary_df,
            output_dir=group_charts_dir,
            suffix="_标准煤(吨标准煤)",
            title_prefix="各能源类型标准煤对比",
        )

        logger.info(f"分组 [{group_name}] 图表生成完成。")
        print(f"图表已保存至: {group_charts_dir}")

        # 显示全年关键指标
        if "日期区间" in summary_df.columns:
            annual_row = summary_df[summary_df["日期区间"] == "全年合计"]
            if not annual_row.empty:
                total_coal = annual_row["总标准煤(吨标准煤)"].values[0]
                efficiency_col = "万元营收标准煤耗(kg/万元营收)"
                efficiency = (
                    annual_row[efficiency_col].values[0]
                    if efficiency_col in annual_row.columns
                    else 0
                )

                print(f"\n>>> [{group_name}] 全年关键指标摘要:")
                print(f"    - 全年总标准煤量: {total_coal:.2f} 吨")
                print(f"    - 万元营收标准煤耗: {efficiency:.2f} kg/万元")

                # 计算单位面积指标
                total_area_config = config.get("total_area", {})
                if total_area_config:
                    total_cost = annual_row["总费用(元)"].values[0]
                    print("    - 单位面积能耗指标:")
                    if isinstance(total_area_config, dict):
                        for area_name, area_val in total_area_config.items():
                            if area_val > 0:
                                cost_per_sqm = total_cost / area_val
                                coal_per_sqm = (total_coal * 1000.0) / area_val
                                print(f"      * {area_name} ({area_val} ㎡):")
                                print(
                                    f"        - 单位面积费用: {cost_per_sqm:.2f} 元/㎡"
                                )
                                print(
                                    f"        - 单位面积标煤: {coal_per_sqm:.4f} kg/㎡"
                                )
                    elif (
                        isinstance(total_area_config, (int, float))
                        and total_area_config > 0
                    ):
                        cost_per_sqm = total_cost / total_area_config
                        coal_per_sqm = (total_coal * 1000.0) / total_area_config
                        print(f"      - 单位面积费用: {cost_per_sqm:.2f} 元/㎡")
                        print(f"      - 单位面积标煤: {coal_per_sqm:.4f} kg/㎡")
            else:
                logger.info(
                    "分组 [%s] 未满足完整全年统计条件，已跳过全年统计。", group_name
                )
                print(
                    f">>> [{group_name}] 未满足完整全年统计条件"
                    "（存在缺失月份或某月整行数据为0），已跳过全年统计。"
                )

    print("\n" + "=" * 50)
    print("能耗费用数据处理工作流执行完成！")
    print("=" * 50)


if __name__ == "__main__":
    main()
