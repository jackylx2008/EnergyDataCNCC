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
import re
import shutil
from typing import Any, cast

import pandas as pd
from core.logging_config import setup_logger
from core.process_energy_data import (
    _parse_env_file,
    _get_revenue_value,
    load_config,
    process_multi_year_comparison,
)
from core.generate_charts import (
    generate_multi_energy_comparison_bar,
    generate_comparison_pie_charts,
)


def _to_float(value: Any) -> float | None:
    """Convert scalar-like values to float when possible."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _get_first_float_from_df(row_df: pd.DataFrame, column: str) -> float | None:
    """Return the first numeric-like value from a DataFrame slice."""
    if row_df.empty or column not in row_df.columns:
        return None

    return _to_float(row_df[column].iloc[0])


def _sum_revenue_values(revenue_data: dict[Any, Any], keys: list[Any]) -> float | None:
    """Sum revenue items only when all referenced values are valid positive numbers."""
    values: list[float] = []
    for key in keys:
        revenue_value = _to_float(revenue_data.get(key))
        if revenue_value is None or revenue_value <= 0:
            return None
        values.append(revenue_value)
    return sum(values)


def _get_revenue_data_for_year(
    revenue_config: dict[Any, Any], year: object
) -> Any:
    """Resolve yearly revenue config using exact key, string key, or integer-like string."""
    revenue_data = revenue_config.get(year)
    if revenue_data is not None:
        return revenue_data

    year_text = str(year).strip()
    if not year_text:
        return None

    revenue_data = revenue_config.get(year_text)
    if revenue_data is not None:
        return revenue_data

    if year_text.isdigit():
        return revenue_config.get(int(year_text))

    return None


def build_year_file_map(config: dict[str, Any]) -> dict[str, str]:
    """
    汇总年度对比文件配置。
    优先使用 config.yaml 中显式声明的 files，同时补齐当前 profile 的
    .env 中 YEAR_COMPARISON_FILE_YYYY 形式的配置。
    """
    comparison_config = cast(dict[str, Any], config.get("year_comparison", {}))
    files_config = comparison_config.get("files", {})
    if not isinstance(files_config, dict):
        files_config = {}
    year_file_map = {
        str(year): str(path).strip()
        for year, path in files_config.items()
        if str(path).strip()
    }

    runtime_config = cast(dict[str, Any], config.get("runtime", {}))
    profile = runtime_config.get("profile", "B25B26")
    env_files = runtime_config.get("env_files", {})
    env_file_map = env_files if isinstance(env_files, dict) else {}
    env_path = str(env_file_map.get(profile, "common.b25b26.env"))

    env_values = _parse_env_file(env_path)
    for key, value in env_values.items():
        match = re.fullmatch(r"YEAR_COMPARISON_FILE_(\d{4})", key)
        if not match:
            continue

        file_path = str(value).strip()
        if not file_path:
            continue

        year = match.group(1)
        year_file_map.setdefault(year, file_path)

    return dict(sorted(year_file_map.items()))


def build_period_efficiency_comparison(
    df: pd.DataFrame, revenue_config: dict[Any, Any], months: list[int] | None = None
) -> pd.DataFrame:
    """
    计算各年份指定月份口径的万元营收标准煤耗对比。
    """
    if df is None or df.empty or "年份" not in df.columns:
        return pd.DataFrame()

    if months is None:
        months = [1, 2, 3]
    months = [int(month) for month in months]

    def extract_month(label):
        match = re.search(r"(\d+)", str(label))
        return int(match.group(1)) if match else None

    month_df = df.copy()
    month_df = month_df[~month_df["日期区间"].str.contains("合计|总计", na=False)].copy()
    month_df["月份"] = month_df["日期区间"].apply(extract_month)
    month_df = month_df[month_df["月份"].isin(months)]

    if month_df.empty:
        return pd.DataFrame()

    period_label = ",".join(f"{month}月" for month in months)
    records = []
    for year, year_df in month_df.groupby("年份"):
        revenue_data = _get_revenue_data_for_year(revenue_config, year)
        if not revenue_data:
            continue

        total_coal = year_df["标准煤(吨标准煤)"].sum()
        period_revenue = None

        if isinstance(revenue_data, dict):
            if all(month in revenue_data for month in months):
                period_revenue = _sum_revenue_values(revenue_data, months)
            elif all(str(month) in revenue_data for month in months):
                period_revenue = _sum_revenue_values(
                    revenue_data, [str(month) for month in months]
                )
            elif all(f"{month}月" in revenue_data for month in months):
                period_revenue = _sum_revenue_values(
                    revenue_data, [f"{month}月" for month in months]
                )

        if not period_revenue or period_revenue <= 0:
            continue

        records.append(
            {
                "日期区间": period_label,
                "年份": year,
                "万元营收标准煤耗(kg/万元营收)": round(
                    total_coal * 1000.0 / period_revenue, 2
                ),
            }
        )

    return pd.DataFrame(records)


def build_monthly_efficiency_markdown_table(
    df: pd.DataFrame, revenue_config: dict[Any, Any], months: list[int] | None = None
) -> str:
    """
    生成按月展示、每三个月插入季度汇总的 Markdown 表格。
    表格增加标准煤消耗量列，并在不同年份数据块之间插入空行。
    """
    if df is None or df.empty or "年份" not in df.columns:
        return ""

    if months is None:
        months = [1, 2, 3]
    months = [int(month) for month in months]

    def extract_month(label):
        match = re.search(r"(\d+)", str(label))
        return int(match.group(1)) if match else None

    month_df = df.copy()
    month_df = month_df[~month_df["日期区间"].str.contains("合计|总计", na=False)].copy()
    month_df["月份"] = month_df["日期区间"].apply(extract_month)
    month_df = month_df[month_df["月份"].isin(months)]
    if month_df.empty:
        return ""

    monthly_coal_df = (
        month_df.groupby(["年份", "月份"], as_index=False)["标准煤(吨标准煤)"].sum()
    )

    rows = []
    quarter_labels = ["一季度", "二季度", "三季度", "四季度"]

    for year in sorted(monthly_coal_df["年份"].unique()):
        revenue_data = _get_revenue_data_for_year(revenue_config, year)
        if not isinstance(revenue_data, dict):
            continue

        year_rows = monthly_coal_df[monthly_coal_df["年份"] == year]
        quarter_buffer = []

        for month in months:
            month_row = year_rows[year_rows["月份"] == month]
            coal = _get_first_float_from_df(month_row, "标准煤(吨标准煤)") or 0.0
            revenue = _to_float(_get_revenue_value(revenue_data, f"{month}月"))
            efficiency = (
                round(coal * 1000.0 / revenue, 2)
                if revenue is not None and revenue > 0
                else None
            )

            rows.append(
                {
                    "年份": year,
                    "期间": f"{month}月",
                    "标准煤消耗量(吨标准煤)": coal,
                    "营收(万元)": revenue,
                    "万元营收标准煤耗(kg/万元营收)": efficiency,
                }
            )
            quarter_buffer.append((month, coal, revenue))

            if len(quarter_buffer) == 3:
                quarter_index = (month - 1) // 3
                quarter_label = (
                    quarter_labels[quarter_index]
                    if 0 <= quarter_index < len(quarter_labels)
                    else f"{quarter_buffer[0][0]}-{quarter_buffer[-1][0]}月汇总"
                )
                quarter_revenue = sum(
                    revenue_item
                    for _, _, revenue_item in quarter_buffer
                    if revenue_item is not None and revenue_item > 0
                )
                has_all_revenue = all(
                    revenue_item is not None and revenue_item > 0
                    for _, _, revenue_item in quarter_buffer
                )
                quarter_coal = sum(coal_item for _, coal_item, _ in quarter_buffer)
                quarter_efficiency = (
                    round(quarter_coal * 1000.0 / quarter_revenue, 2)
                    if has_all_revenue and quarter_revenue > 0
                    else None
                )
                rows.append(
                    {
                        "年份": year,
                        "期间": quarter_label,
                        "标准煤消耗量(吨标准煤)": quarter_coal,
                        "营收(万元)": quarter_revenue if has_all_revenue else None,
                        "万元营收标准煤耗(kg/万元营收)": quarter_efficiency,
                    }
                )
                quarter_buffer = []

    if not rows:
        return ""

    table_df = pd.DataFrame(rows)
    table_df["标准煤消耗量(吨标准煤)"] = table_df["标准煤消耗量(吨标准煤)"].apply(
        lambda value: f"{value:.2f}" if pd.notna(value) else "-"
    )
    table_df["营收(万元)"] = table_df["营收(万元)"].apply(
        lambda value: f"{value:.2f}" if pd.notna(value) else "-"
    )
    table_df["万元营收标准煤耗(kg/万元营收)"] = table_df[
        "万元营收标准煤耗(kg/万元营收)"
    ].apply(lambda value: f"{value:.2f}" if pd.notna(value) else "-")

    headers = [
        "年份",
        "期间",
        "标准煤消耗量(吨标准煤)",
        "营收(万元)",
        "万元营收标准煤耗(kg/万元营收)",
    ]
    lines = [
        "| " + " | ".join(headers) + " |",
        "| :---: | :---: | :---: | :---: | :---: |",
    ]
    previous_year = None
    for _, row in table_df.iterrows():
        if previous_year is not None and row["年份"] != previous_year:
            lines.append("")
        lines.append(
            "| "
            + " | ".join(str(row[header]) for header in headers)
            + " |"
        )
        previous_year = row["年份"]

    return "\n".join(lines)


def build_area_efficiency_comparison(
    df: pd.DataFrame, total_area_config: Any
) -> pd.DataFrame:
    """
    计算各年份单位面积标准煤和单位面积能耗费用对比。
    基于当前对比数据口径汇总。
    """
    if df is None or df.empty or "年份" not in df.columns:
        return pd.DataFrame()

    if isinstance(total_area_config, dict):
        area_value = next(
            (
                float(value)
                for value in total_area_config.values()
                if isinstance(value, (int, float)) and value > 0
            ),
            None,
        )
    elif isinstance(total_area_config, (int, float)) and total_area_config > 0:
        area_value = float(total_area_config)
    else:
        area_value = None

    if not area_value:
        return pd.DataFrame()

    plot_df = df[~df["日期区间"].str.contains("合计|总计", na=False)].copy()
    if plot_df.empty:
        return pd.DataFrame()

    summary_df = (
        plot_df.groupby("年份", as_index=False)[["费用(元)", "标准煤(吨标准煤)"]].sum()
    )
    summary_df["单位平米能耗费用(元/㎡)"] = (
        summary_df["费用(元)"] / area_value
    ).round(4)
    summary_df["单位平米标准煤(kg/㎡)"] = (
        summary_df["标准煤(吨标准煤)"] * 1000.0 / area_value
    ).round(4)

    return summary_df[
        ["年份", "单位平米标准煤(kg/㎡)", "单位平米能耗费用(元/㎡)"]
    ]


def main() -> None:
    # 设置对比工作流专用日志
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/main_comparison.log")
    logger.info("开始执行 [年度对比] 数据处理工作流...")

    # 加载配置
    config = cast(dict[str, Any], load_config())
    comparison_config = cast(dict[str, Any], config.get("year_comparison", {}))
    year_file_map = build_year_file_map(config)
    operating_revenue_config = cast(
        dict[Any, Any], config.get("operating_revenue", {})
    )
    total_area_config = config.get("total_area", {})
    paths_config = cast(dict[str, Any], config.get("paths", {}))
    output_dir = str(paths_config.get("output_dir", ""))
    comparison_output_dir = os.path.join(output_dir, "charts_年度对比")

    if not output_dir:
        logger.error("配置缺少 paths.output_dir。")
        print("Error: 配置缺少 paths.output_dir，请检查 config.yaml。")
        return

    if not year_file_map:
        logger.error("未找到任何年度对比文件配置。")
        print("Error: 未找到年度对比文件配置，请检查 config.yaml 或当前 profile 的 .env。")
        return

    missing_year_file_map = {
        year: file_path
        for year, file_path in year_file_map.items()
        if not os.path.exists(file_path)
    }
    valid_year_file_map = {
        year: file_path
        for year, file_path in year_file_map.items()
        if os.path.exists(file_path)
    }

    for year, file_path in missing_year_file_map.items():
        logger.warning(f"年份 {year} 的文件不存在，已跳过: {file_path}")

    if not valid_year_file_map:
        logger.error("所有已配置的年度对比文件都不存在。")
        print("未找到可读取的年度对比文件，请检查以下路径配置：")
        for year, file_path in year_file_map.items():
            print(f"  - {year}: {file_path}")
        return

    if len(valid_year_file_map) < 2:
        logger.warning(
            f"仅找到 {len(valid_year_file_map)} 个有效年份，将按单年份生成分析结果: "
            f"{', '.join(valid_year_file_map.keys())}"
        )

    logger.info(
        f"本次纳入年度分析的年份: {', '.join(valid_year_file_map.keys())}"
    )

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

    # 1. 数据处理：汇总之不同年份的数据，并计算标准煤
    print("正在处理多年份数据...")
    df_combined = process_multi_year_comparison(valid_year_file_map, output_dir)

    if df_combined is None or df_combined.empty:
        logger.warning("处理后未获取到有效对比数据。")
        print("未获取到有效的对比数据，请检查输入文件路径和格式。")
        return

    # 根据 config.yaml 中的 months 进行过滤
    raw_comparison_months = comparison_config.get("months", [])
    comparison_months = (
        [int(month) for month in raw_comparison_months]
        if isinstance(raw_comparison_months, list)
        else []
    )
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

    period_efficiency_df = build_period_efficiency_comparison(
        df=df_combined,
        revenue_config=operating_revenue_config,
        months=comparison_months or [1, 2, 3],
    )
    period_label = ",".join(f"{month}月" for month in (comparison_months or [1, 2, 3]))
    markdown_table = build_monthly_efficiency_markdown_table(
        df=df_combined,
        revenue_config=operating_revenue_config,
        months=comparison_months or [1, 2, 3],
    )
    if markdown_table:
        print(f"\n>>> {period_label}万元营收标准煤耗对比（Markdown表格）:")
        print(markdown_table)
    elif period_efficiency_df.empty:
        logger.warning("未能生成指定月份万元营收标准煤耗对比数据。")
        print(f"\n>>> {period_label}万元营收标准煤耗对比: 未获取到可用数据")

    area_efficiency_df = build_area_efficiency_comparison(
        df=df_combined,
        total_area_config=total_area_config,
    )
    if not area_efficiency_df.empty:
        print("\n>>> 截至3月累计单位平米指标对比:")
        for _, row in area_efficiency_df.sort_values("年份").iterrows():
            year_label = str(row["年份"])
            coal_per_sqm = _to_float(row["单位平米标准煤(kg/㎡)"]) or 0.0
            cost_per_sqm = _to_float(row["单位平米能耗费用(元/㎡)"]) or 0.0
            print(
                f"    - {year_label}年截至3月累计: "
                f"单位平米标准煤 {coal_per_sqm:.4f} kg/㎡, "
                f"单位平米能耗费用 {cost_per_sqm:.4f} 元/㎡"
            )
    else:
        logger.warning("未能生成截至3月累计单位平米指标对比数据。")
        print("\n>>> 截至3月累计单位平米指标对比: 未获取到可用数据")

    print()

    # B. 记录日志完成
    logger.info("年度对比工作流执行完成。")
    print("\n--- 年度对比处理完成 ---")
    print(f"对比数据明细: {os.path.join(output_dir, 'energy_comparison_details.xlsx')}")
    print(f"对比分析图表: {comparison_output_dir}")


if __name__ == "__main__":
    main()
