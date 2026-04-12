"""
图表生成模块
==========

此脚本用于读取处理后的能源汇总数据 (Excel)，并生成可视化的统计图表。
生成的图表保存在 output/charts 目录下。

包含的图表类型:
1. 饼图 (Pie Chart): 展示每个时间区间的能源费用分布占比。
2. 堆叠柱状图 (Stacked Bar Chart): 展示不同时间区间的总费用及各能源类型的构成。
3. 分组柱状图 (Grouped Bar Chart): 并排展示不同时间区间各能源类型的费用对比。
"""

import logging
import os
import colorsys

import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import pandas as pd
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter

try:
    import yaml
except ImportError:  # pragma: no cover - PyYAML optional
    yaml = None

from logging_config import setup_logger

# Setup general chart module logger
CHART_LOGGER = setup_logger(
    log_level=logging.INFO, log_file="./logs/charts.log", filemode="a"
)

# Setup a dedicated logger for energy data details
# We manually add a handler to avoid clearing root handlers via setup_logger
DATA_LOGGER = logging.getLogger("energy_data_details")
DATA_LOGGER.setLevel(logging.INFO)
DATA_LOGGER.propagate = (
    False  # Don't send to root logger to avoid duplication in charts.log
)
_data_handler = logging.FileHandler(
    "./logs/energy_data_details.log", mode="a", encoding="utf-8"
)
_data_handler.setFormatter(logging.Formatter("%(asctime)s - %(message)s"))
DATA_LOGGER.addHandler(_data_handler)

# Configure Chinese font support for Matplotlib
plt.rcParams["font.sans-serif"] = [
    "SimHei",
    "Microsoft YaHei",
    "Arial Unicode MS",
]  # Windows/Mac compatible
plt.rcParams["axes.unicode_minus"] = False

DEFAULT_ENERGY_COLOR_MAP = {
    "电": "#A8C8E1",
    "采暖用热": "#C1E0A5",
    "生活热水用热": "#F2B3B1",
    "自来水": "#F28C28",
    "中水": "#7E5AA7",
    "燃气": "#B5754C",
}


def load_energy_color_map(config_path: str = "config.yaml") -> dict:
    """Load palette from config file with sane fallbacks."""

    colors = DEFAULT_ENERGY_COLOR_MAP.copy()
    if yaml is None:
        logging.getLogger(__name__).warning("未安装 PyYAML，使用默认配色")
        return colors

    if not os.path.exists(config_path):
        return colors

    try:
        with open(config_path, "r", encoding="utf-8") as cfg_file:
            config = yaml.safe_load(cfg_file) or {}
    except Exception as exc:  # pragma: no cover - defensive
        logging.getLogger(__name__).warning("读取配色配置失败，使用默认值: %s", exc)
        return colors

    palette = config.get("colors", {}).get("energy", {})
    if isinstance(palette, dict):
        sanitized = {
            str(key): str(value)
            for key, value in palette.items()
            if isinstance(value, str)
        }
        if sanitized:
            colors.update(sanitized)

    return colors


ENERGY_COLOR_MAP = load_energy_color_map()


def get_color_sequence(labels):
    """Return a color list aligned with known energy type order."""

    default_colors = list(ENERGY_COLOR_MAP.values()) or ["#999999"]
    colors = []
    fallback_index = 0
    for label in labels:
        if label in ENERGY_COLOR_MAP:
            colors.append(ENERGY_COLOR_MAP[label])
        else:
            colors.append(default_colors[fallback_index % len(default_colors)])
            fallback_index += 1
    return colors


def lighten_color(color, amount=0.5):
    """
    极简轻量化颜色。amount 越接近 1 则越淡。
    """
    try:
        rgb = mcolors.to_rgb(color)
        hls = colorsys.rgb_to_hls(*rgb)
        # 增加亮度 (Lightness)
        new_l = 1 - amount * (1 - hls[1])
        new_rgb = colorsys.hls_to_rgb(hls[0], new_l, hls[2])
        return mcolors.to_hex(new_rgb)
    except:
        return color


def generate_pie_charts(
    input_file: str | pd.DataFrame = "./output/energy_usage_summary.xlsx",
    output_dir: str = "./output/charts",
    suffix: str = "_费用(元)",
    title_prefix: str = "能源费用分布",
):
    """
    生成费用/用量分布饼图。

    遍历汇总数据中的每一行 (每个日期区间)，为每个区间生成一个饼图，
    显示不同能源类型的费用占比。

    参数:
        input_file: Excel 文件路径或包含汇总数据的 DataFrame。
        output_dir: 图表输出目录。
        suffix: 匹配的列后缀 (例如 "_费用(元)" 或 "_标准煤(吨标准煤)")。
        title_prefix: 标题前缀。
    """
    # Setup logging
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    if not isinstance(input_file, pd.DataFrame):
        if not os.path.exists(input_file):
            logger.error(f"未找到输入文件: {input_file}")
            return
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            logger.error(f"读取 Excel 失败: {e}")
            return
    else:
        df = input_file

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Identify relevant columns
        cost_cols = [col for col in df.columns if col.endswith(suffix)]
        total_col = "总费用(元)" if suffix == "_费用(元)" else f"总{suffix.lstrip('_')}"

        if not cost_cols:
            logger.warning(f"未找到后缀为 {suffix} 的列。")
            return

        for index, row in df.iterrows():
            date_range = row["日期区间"]

            # Extract data for this row
            values = []
            labels = []
            log_parts = [f"--- {title_prefix} 详情 ({date_range}) ---"]

            for col in cost_cols:
                val = row[col]
                # 仅包含数值大于 0 的项
                if val > 0 and col != total_col:
                    values.append(val)
                    # Example: "电_费用(元)" -> "电"
                    energy_type = col.replace(suffix, "")
                    labels.append(energy_type)

                    # Log physical usage if available
                    usage_col = f"{energy_type}_用量"
                    usage_str = ""
                    if usage_col in df.columns:
                        usage_val = row[usage_col]
                        usage_str = f" (用量: {usage_val})"

                    unit_fmt = ",.2f"  # All use 2 decimals now
                    log_parts.append(f"{energy_type}: {val:{unit_fmt}}{usage_str}")

            if not values or sum(values) <= 0:
                logger.info(f"{date_range} 无数据或总量为0，跳过图表生成。")
                continue

            # Log to dedicated data logger
            DATA_LOGGER.info(" | ".join(log_parts))

            total_cost = sum(values)

            # Create Pie Chart
            # Large canvas keeps legend readable
            plt.figure(figsize=(16, 10))

            # Pie chart
            # We use a legend to avoid label overlap on the chart itself
            # 只有大于 0 的项才显示百分比标注，避免 0% 堆叠
            pie_colors = get_color_sequence(labels)
            wedges, texts, autotexts = plt.pie(  # type: ignore
                values,
                colors=pie_colors,
                autopct=lambda p: f"{p:.1f}%" if p > 0 else "",
                startangle=140,
                pctdistance=0.75,
                textprops={"fontsize": 18},
            )

            # Adjust distance for small slices (0 < percentage < 5%) to avoid overlap
            small_slice_idx = 0
            for i, autotext in enumerate(autotexts):
                percentage = (values[i] / total_cost * 100) if total_cost > 0 else 0
                if 0 < percentage < 5:
                    # Alternating distances: 0.65 and 0.88 (original was 0.75)
                    new_dist = 0.65 if small_slice_idx % 2 == 0 else 0.88
                    curr_x, curr_y = autotext.get_position()
                    # Scale based on original pctdistance of 0.75
                    scale = new_dist / 0.75
                    autotext.set_position((curr_x * scale, curr_y * scale))
                    # Slightly smaller font for very small slices
                    if percentage < 2:
                        autotext.set_fontsize(14)
                    small_slice_idx += 1

            plt.title(
                f"{title_prefix} - {date_range}", fontsize=28, pad=20
            )  # Increased title font size
            plt.axis(
                "equal"
            )  # Equal aspect ratio ensures that pie is drawn as a circle.

            unit = suffix.split("(_")[-1].strip(")") if "(" in suffix else "单位"
            if suffix == "_费用(元)":
                unit = "元"
            elif "标准煤" in suffix:
                unit = "吨标准煤"

            # Create detailed legend labels
            val_fmt = ",.2f"  # All use 2 decimals now (Cost and Tons)
            legend_labels = [
                f"{label}: {val:{val_fmt}}{unit}" for label, val in zip(labels, values)
            ]

            # Add legend to the right
            # Use fixed bbox_to_anchor to prevent legend size from affecting pie size
            # bbox_transform=plt.gcf().transFigure ensures coordinates are relative to the whole figure (0 to 1)
            # Pie Right Edge is approx at X=0.53 (calculated from subplots_adjust and equal axis)
            plt.legend(
                wedges,
                legend_labels,
                title="分项明细",
                loc="center left",
                bbox_to_anchor=(0.53, 0.51),
                bbox_transform=plt.gcf().transFigure,
                fontsize=30,
                title_fontsize=30,
            )

            # Add total cost at the bottom
            total_label = "总额" if suffix != "_费用(元)" else "总费用"
            plt.figtext(
                0.3,
                0.05,
                f"{total_label}: {total_cost:{val_fmt}} {unit}",
                ha="center",
                fontsize=26,
                fontweight="bold",
                color="#333333",
            )

            # Use subplots_adjust instead of tight_layout to ensure fixed pie chart size
            # regardless of legend items count or label lengths.
            # Shifted left (right=0.55) to accommodate longer legend labels
            plt.subplots_adjust(left=0.02, bottom=0.12, right=0.55, top=0.9)

            # Save chart
            # Clean filename
            allowed_chars = {" ", ".", "-", "_"}
            clean_chars = [
                c for c in str(date_range) if c.isalnum() or c in allowed_chars
            ]
            safe_date_range = "".join(clean_chars).strip()
            file_prefix = "cost" if "费用" in title_prefix else "coal"
            output_path = os.path.join(
                output_dir, f"{file_prefix}_distribution_{safe_date_range}.png"
            )

            plt.savefig(output_path)
            plt.close()

            logger.info(f"已生成图表: {output_path}")
            print(f"已生成图表: {output_path}")

        # 生成全年汇总饼图
        generate_annual_pie_chart(input_file, output_dir, suffix, title_prefix)

    except Exception as e:
        logger.error(f"生成图表失败: {e}", exc_info=True)


def generate_annual_pie_chart(
    input_file: str | pd.DataFrame = "./output/energy_usage_summary.xlsx",
    output_dir: str = "./output/charts",
    suffix: str = "_费用(元)",
    title_prefix: str = "能源费用分布",
):
    """
    生成全年能源费用/用量分布饼图。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    try:
        if not isinstance(input_file, pd.DataFrame):
            df = pd.read_excel(input_file)
        else:
            df = input_file

        cost_cols = [col for col in df.columns if col.endswith(suffix)]
        total_col = "总费用(元)" if suffix == "_费用(元)" else f"总{suffix.lstrip('_')}"

        values = []
        labels = []
        log_parts = [f"===全年 {title_prefix} 汇总 ==="]

        # 过滤掉 "全年合计" 行，避免重复计算
        df_monthly = df[df["日期区间"] != "全年合计"]

        for col in cost_cols:
            val = df_monthly[col].sum()
            # 仅包含总额大于 0 的项
            if val > 0 and col != total_col:
                values.append(val)
                energy_type = col.replace(suffix, "")
                labels.append(energy_type)

                # Log physical usage if available
                usage_col = f"{energy_type}_用量"
                usage_str = ""
                if usage_col in df_monthly.columns:
                    usage_val = df_monthly[usage_col].sum()
                    usage_str = f" (用量: {usage_val})"

                unit_fmt = ",.2f"  # All use 2 decimals now
                log_parts.append(f"{energy_type}: {val:{unit_fmt}}{usage_str}")

        if not values or sum(values) <= 0:
            return

        # Log to dedicated data logger
        DATA_LOGGER.info(" | ".join(log_parts))

        total_cost = sum(values)

        plt.figure(figsize=(16, 10))
        pie_colors = get_color_sequence(labels)
        wedges, texts, autotexts = plt.pie(  # type: ignore
            values,
            colors=pie_colors,
            autopct=lambda p: f"{p:.1f}%" if p > 0 else "",
            startangle=140,
            pctdistance=0.75,
            textprops={"fontsize": 18},
        )

        # 交错显示小占比标注
        small_slice_idx = 0
        for i, autotext in enumerate(autotexts):
            percentage = (values[i] / total_cost * 100) if total_cost > 0 else 0
            if 0 < percentage < 5:
                # 交错布局：0.65 和 0.88
                new_dist = 0.65 if small_slice_idx % 2 == 0 else 0.88
                curr_x, curr_y = autotext.get_position()
                scale = new_dist / 0.75
                autotext.set_position((curr_x * scale, curr_y * scale))
                small_slice_idx += 1
                if percentage < 2:
                    autotext.set_fontsize(14)
                small_slice_idx += 1

        plt.title(f"全年{title_prefix}汇总", fontsize=28, pad=20)
        plt.axis("equal")

        unit = "元" if suffix == "_费用(元)" else "吨标准煤"
        val_fmt = ",.2f"  # All use 2 decimals now
        legend_labels = [
            f"{label}: {val:{val_fmt}}{unit}" for label, val in zip(labels, values)
        ]
        plt.legend(
            wedges,
            legend_labels,
            title="分项明细",
            loc="center left",
            bbox_to_anchor=(0.53, 0.51),
            bbox_transform=plt.gcf().transFigure,
            fontsize=30,
            title_fontsize=30,
        )

        plt.figtext(
            0.3,
            0.05,
            f"全年总额: {total_cost:{val_fmt}} {unit}",
            ha="center",
            fontsize=26,
            fontweight="bold",
            color="#333333",
        )

        plt.subplots_adjust(left=0.02, bottom=0.12, right=0.55, top=0.9)
        file_prefix = "cost" if "费用" in title_prefix else "coal"
        output_path = os.path.join(
            output_dir, f"{file_prefix}_distribution_annual_summary.png"
        )
        plt.savefig(output_path)
        plt.close()

        logger.info(f"已生成全年汇总饼图: {output_path}")
        print(f"已生成全年汇总饼图: {output_path}")

    except Exception as e:
        logger.error(f"生成全年汇总饼图失败: {e}", exc_info=True)


def generate_energy_type_distribution_bar(
    input_file: str | pd.DataFrame = "./output/energy_usage_summary.xlsx",
    output_dir: str = "./output/charts",
    suffix: str = "_费用(元)",
    title_prefix: str = "各能源类型费用对比",
):
    """
    生成能源类型分布柱状图。
    X 轴为能源类型 (电, 采暖用热, 生活热水用热等)，Y 轴为数值 (费用或标准煤)。
    展示的是全年的累计数值。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    try:
        if not isinstance(input_file, pd.DataFrame):
            df = pd.read_excel(input_file)
        else:
            df = input_file

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 1. 提取数据 (优先使用 全年合计 行)
        df_annual = df[df["日期区间"] == "全年合计"]
        if df_annual.empty:
            # 如果没有合计行，则手动汇总
            df_monthly = df[df["日期区间"] != "全年合计"]
            totals = df_monthly.sum(numeric_only=True)
        else:
            totals = df_annual.iloc[0]

        # 2. 识别相关列
        cost_cols = [
            col for col in df.columns if col.endswith(suffix) and "总" not in col
        ]

        data = {}
        for col in cost_cols:
            val = totals[col]
            if val > 0:
                label = col.replace(suffix, "")
                data[label] = val

        if not data:
            logger.warning(f"未找到后缀为 {suffix} 的有效数据，跳过柱状图。")
            return

        # 3. 排序以便展示 (按数值从大到小)
        sorted_data = dict(sorted(data.items(), key=lambda item: item[1], reverse=True))
        labels = list(sorted_data.keys())
        values = list(sorted_data.values())

        # 4. 绘图
        plt.figure(figsize=(16, 12))
        ax = plt.gca()

        bar_colors = get_color_sequence(labels)
        bars = ax.bar(labels, values, color=bar_colors, alpha=0.9, width=0.6)

        plt.title(f"全年{title_prefix}", fontsize=30, pad=30, fontweight="bold")

        # 处理 Y 轴单位
        unit = "元"
        use_wan_yuan = False
        if suffix == "_费用(元)":
            unit = "元"
            if max(values) > 50000:
                use_wan_yuan = True
                unit = "万元"
        elif "标准煤" in suffix:
            unit = "吨标准煤"

        if use_wan_yuan:
            ax.yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x/10000:.1f}"))
            plt.ylabel(f"费用 ({unit})", fontsize=24, labelpad=15)
        else:
            ax.yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x:,.0f}"))
            plt.ylabel(f"数值 ({unit})", fontsize=24, labelpad=15)

        plt.xticks(fontsize=22)
        plt.yticks(fontsize=20)

        # 添加数值标签
        for bar in bars:
            height = bar.get_height()
            label_val = f"{height/10000:,.1f}" if use_wan_yuan else f"{height:,.0f}"
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height,
                label_val,
                ha="center",
                va="bottom",
                fontsize=18,
                fontweight="bold",
            )

        plt.grid(axis="y", linestyle="--", alpha=0.7)
        plt.tight_layout()

        # 保存
        file_prefix = "cost_type" if "费用" in title_prefix else "coal_type"
        output_path = os.path.join(output_dir, f"{file_prefix}_bar_annual.png")
        plt.savefig(output_path)
        plt.close()

        logger.info(f"已生成能源类型分布柱状图: {output_path}")
        print(f"已生成能源类型分布柱状图: {output_path}")

    except Exception as e:
        logger.error(f"生成能源类型柱状图失败: {e}", exc_info=True)


def generate_cost_bar_chart(
    input_file: str | pd.DataFrame = "./output/energy_usage_summary.xlsx",
    output_dir: str = "./output/charts",
):
    """
    生成费用对比堆叠柱状图。

    以日期区间为 X 轴，费用为 Y 轴，展示各区间的总费用。
    不同能源类型的费用在柱状图中堆叠显示，方便比较总费用及构成。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    if not isinstance(input_file, pd.DataFrame):
        if not os.path.exists(input_file):
            logger.error(f"未找到输入文件: {input_file}")
            return
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            logger.error(f"读取 Excel 失败: {e}")
            return
    else:
        df = input_file

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Filter cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]
        if not cost_cols:
            logger.warning("未找到用于柱状图的费用列。")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_费用(元)", "") for col in cost_cols}

        # 过滤掉 "全年合计" 行，避免比例失调
        df_plot = df[df["日期区间"] != "全年合计"]
        plot_df = df_plot.set_index("日期区间")[cost_cols].rename(columns=rename_map)

        # Filter out rows with 0 total cost
        plot_df = plot_df[plot_df.sum(axis=1) > 0]

        if plot_df.empty:
            logger.info("无有效的柱状图数据。")
            return

        # Create figure
        plt.figure(figsize=(18, 12))
        ax = plt.gca()

        # Plot stacked bar chart with shared color palette
        bar_colors = get_color_sequence(plot_df.columns.tolist())
        plot_df.plot(
            kind="bar",
            stacked=True,
            ax=ax,
            width=0.6,
            alpha=0.9,
            color=bar_colors,
        )

        # Styling
        plt.title("各区间能源费用对比", fontsize=30, pad=25)
        plt.xlabel("日期区间", fontsize=24, labelpad=15)
        plt.ylabel("费用 (万元)", fontsize=24, labelpad=15)
        ax.yaxis.set_major_formatter(
            FuncFormatter(lambda value, _: f"{value / 10000:.1f}")
        )
        plt.xticks(rotation=0, fontsize=20)
        plt.yticks(fontsize=20)

        # Legend
        plt.legend(
            title="能源类型",
            fontsize=18,
            title_fontsize=20,
            bbox_to_anchor=(1.01, 1),
            loc="upper left",
        )

        # Add total labels on top of bars
        totals = plot_df.sum(axis=1)
        for i, total in enumerate(totals):
            ax.text(
                i,
                total,
                f"{total:,.0f}",
                ha="center",
                va="bottom",
                fontsize=18,
                fontweight="bold",
                color="black",
            )

        # Add value labels inside bars (only for significant values)
        for c in ax.containers:
            # Create labels
            labels = []
            for v in c:
                height = v.get_height()
                # Only label if height is > 5% of max total to avoid clutter
                if height > totals.max() * 0.05:
                    labels.append(f"{height:,.0f}")
                else:
                    labels.append("")
            ax.bar_label(
                c,  # type: ignore
                labels=labels,
                label_type="center",
                fontsize=18,
                color="black",
                fontweight="bold",
            )

        plt.tight_layout()

        output_path = os.path.join(output_dir, "cost_comparison_bar.png")
        plt.savefig(output_path)
        plt.close()

        logger.info(f"已生成柱状图: {output_path}")
        print(f"已生成柱状图: {output_path}")

    except Exception as e:
        logger.error(f"生成柱状图失败: {e}", exc_info=True)


def generate_coal_bar_chart(
    input_file: str | pd.DataFrame = "./output/energy_usage_summary.xlsx",
    output_dir: str = "./output/charts",
):
    """
    生成标准煤消耗对比堆叠柱状图。

    以日期区间为 X 轴，标准煤为 Y 轴，展示各区间的总消耗。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    if not isinstance(input_file, pd.DataFrame):
        if not os.path.exists(input_file):
            logger.error(f"未找到输入文件: {input_file}")
            return
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            logger.error(f"读取 Excel 失败: {e}")
            return
    else:
        df = input_file

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Filter coal columns
        coal_cols = [col for col in df.columns if col.endswith("_标准煤(吨标准煤)")]
        if not coal_cols:
            logger.warning("未找到用于柱状图的标准煤列。")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_标准煤(吨标准煤)", "") for col in coal_cols}

        # 过滤掉 "全年合计" 行，避免比例失调
        df_plot = df[df["日期区间"] != "全年合计"]
        plot_df = df_plot.set_index("日期区间")[coal_cols].rename(columns=rename_map)

        # Filter out rows with 0 total
        plot_df = plot_df[plot_df.sum(axis=1) > 0]

        # Filter out columns with 0 total (e.g. Water/Reclaimed Water usually 0 coal)
        plot_df = plot_df.loc[:, plot_df.sum(axis=0) > 0]

        if plot_df.empty:
            logger.info("无有效的柱状图数据。")
            return

        # Create figure
        plt.figure(figsize=(18, 12))
        ax = plt.gca()

        # Plot stacked bar chart with shared color palette
        bar_colors = get_color_sequence(plot_df.columns.tolist())
        plot_df.plot(
            kind="bar",
            stacked=True,
            ax=ax,
            width=0.6,
            alpha=0.9,
            color=bar_colors,
        )

        # Styling
        plt.title("各区间标准煤消耗对比", fontsize=30, pad=25)
        plt.xlabel("日期区间", fontsize=24, labelpad=15)
        plt.ylabel("标准煤 (吨)", fontsize=24, labelpad=15)

        # Y-axis formatter: No scaling, just comma separation
        ax.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}"))
        plt.xticks(rotation=0, fontsize=20)
        plt.yticks(fontsize=20)

        # Legend
        plt.legend(
            title="能源类型",
            fontsize=18,
            title_fontsize=20,
            bbox_to_anchor=(1.01, 1),
            loc="upper left",
        )

        # Add total labels on top of bars
        totals = plot_df.sum(axis=1)
        for i, total in enumerate(totals):
            ax.text(
                i,
                total,
                f"{total:,.2f}",
                ha="center",
                va="bottom",
                fontsize=18,
                fontweight="bold",
                color="black",
            )

        # Add value labels inside bars (only for significant values)
        for c in ax.containers:
            # Create labels
            labels = []
            for v in c:
                height = v.get_height()
                # Only label if height is > 5% of max total to avoid clutter
                if height > totals.max() * 0.05:
                    labels.append(f"{height:,.2f}")
                else:
                    labels.append("")
            ax.bar_label(
                c,  # type: ignore
                labels=labels,
                label_type="center",
                fontsize=18,
                color="black",
                fontweight="bold",
            )

        plt.tight_layout()

        output_path = os.path.join(output_dir, "coal_comparison_bar.png")
        plt.savefig(output_path)
        plt.close()

        logger.info(f"已生成柱状图: {output_path}")
        print(f"已生成柱状图: {output_path}")

    except Exception as e:
        logger.error(f"生成柱状图失败: {e}", exc_info=True)


def generate_grouped_bar_chart(
    input_file: str | pd.DataFrame = "./output/energy_usage_summary.xlsx",
    output_dir: str = "./output/charts",
):
    """
    生成分项费用对比分组柱状图。
    根据费用的数量级，自动将数据分为两组生成两张图表：
    1. 主要能源：电、采暖热费
    2. 其他能源：自来水、燃气、生活热水热费等
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    if not isinstance(input_file, pd.DataFrame):
        if not os.path.exists(input_file):
            logger.error(f"未找到输入文件: {input_file}")
            return
        try:
            df = pd.read_excel(input_file)
        except Exception as e:
            logger.error(f"读取 Excel 失败: {e}")
            return
    else:
        df = input_file

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # Filter cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]
        if not cost_cols:
            logger.warning("未找到用于分组柱状图的费用列。")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_费用(元)", "") for col in cost_cols}

        # 过滤掉 "全年合计" 行，避免比例失调
        df_plot = df[df["日期区间"] != "全年合计"]
        full_plot_df = df_plot.set_index("日期区间")[cost_cols].rename(
            columns=rename_map
        )

        # Remove rows with 0 total cost
        full_plot_df = full_plot_df[full_plot_df.sum(axis=1) > 0]

        if full_plot_df.empty:
            logger.info("无有效的分组柱状图数据。")
            return

        # Define groups
        major_types = ["电", "采暖用热"]

        # Sub-function to generate chart for a subset of columns
        def create_chart_for_columns(columns, suffix, title_suffix):
            valid_cols = [c for c in columns if c in full_plot_df.columns]
            if not valid_cols:
                return

            subset_df = full_plot_df[valid_cols]
            # Skip if all zeros
            if subset_df.sum().sum() == 0:
                return

            # Create figure and primary axis
            fig, ax1 = plt.subplots(figsize=(20, 12))

            # Plot grouped bars on ax1 with consistent palette
            group_colors = get_color_sequence(subset_df.columns.tolist())
            subset_df.plot(
                kind="bar",
                ax=ax1,
                width=0.6,
                alpha=0.9,
                rot=0,
                color=group_colors,
                legend=False,
            )

            # Styling
            ax1.set_title(
                "水、电、热费用统计",
                fontsize=30,
                pad=30,
                fontweight="bold",
            )
            ax1.set_xlabel("", fontsize=30)  # X-axis label is redundant with dates
            ax1.set_ylabel("分项费用 (万元)", fontsize=30, labelpad=15)
            ax1.yaxis.set_major_formatter(
                FuncFormatter(lambda value, _: f"{value / 10000:.1f}")
            )

            # Ticks
            ax1.tick_params(axis="x", rotation=0, labelsize=30)
            ax1.tick_params(axis="y", labelsize=30)

            # Grid (Horizontal only)
            ax1.grid(axis="y", linestyle="--", alpha=0.3)

            # Legend
            legend_labels = subset_df.columns.tolist()
            legend_handles = [
                Patch(facecolor=color, edgecolor=color) for color in group_colors
            ]

            # Place legend at the bottom
            fig.legend(
                legend_handles,
                legend_labels,
                loc="lower center",
                bbox_to_anchor=(0.5, 0.02),
                ncol=len(legend_labels),
                fontsize=25,
                handlelength=2.0,
                frameon=False,
            )

            # Adjust layout to make room for legend and rotated labels
            plt.subplots_adjust(bottom=0.2, top=0.9)

            # Add value labels on top of each bar (ax1)
            for c in ax1.containers:
                labels = []
                for v in c:
                    height = v.get_height()
                    if height > 0:
                        labels.append(f"{height:,.0f}")
                    else:
                        labels.append("")

                ax1.bar_label(
                    c,  # type: ignore
                    labels=labels,
                    label_type="edge",
                    fontsize=20,
                    padding=3,
                    rotation=0,
                )

            output_path = os.path.join(output_dir, f"cost_grouped_bar_{suffix}.png")
            plt.savefig(output_path)
            plt.close()

            logger.info(f"已生成分组柱状图: {output_path}")
            print(f"已生成分组柱状图: {output_path}")

        # 1. Generate Major Chart
        # 主要能源包括：电、采暖用热，由于数值较大通常单独展示
        create_chart_for_columns(major_types, "major", "主要能源")

        # 2. Generate Minor Chart (All others)
        # 其他能源包括：自来水、燃气、生活热水等，数值相对较小
        minor_types = [c for c in full_plot_df.columns if c not in major_types]
        create_chart_for_columns(minor_types, "minor", "其他能源")

    except Exception as e:
        logger.error(f"生成分组柱状图失败: {e}", exc_info=True)


def generate_comparison_bar_chart(
    df: pd.DataFrame,
    output_dir: str,
    value_col: str = "费用(元)",
    title: str = "年度能源对比分析",
    ylabel: str = "数值",
    filename: str = "year_comparison.png",
    use_wan_yuan: bool = False,
):
    """
    生成年度同比对比柱状图。
    X 轴为月份，不同年份为并列柱子。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    if df is None or df.empty:
        return

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 1. 数据准备：按 月份 和 年份 分组汇总
        # 确保月份排序正确（1月, 2月...）
        def get_month_num(m_str):
            import re

            match = re.search(r"(\d+)", str(m_str))
            return int(match.group(1)) if match else 999

        # 过滤掉“全年合计”或“总计”
        plot_df_raw = df[~df["日期区间"].str.contains("合计|总计", na=False)].copy()
        plot_df_raw["_month_num"] = plot_df_raw["日期区间"].apply(get_month_num)

        # Pivot: Index=日期区间, Columns=年份
        pivot_df = plot_df_raw.pivot_table(
            index=["_month_num", "日期区间"],
            columns="年份",
            values=value_col,
            aggfunc="sum",
        )
        # 只保留日期区间作为索引
        pivot_df.index = pivot_df.index.get_level_values("日期区间")

        if pivot_df.empty:
            logger.warning(f"对比图表 {filename} 无有效数据。")
            return

        # 2. 绘图
        plt.figure(figsize=(20, 12))
        ax = plt.gca()

        # 使用内置配色或循环色
        pivot_df.plot(kind="bar", ax=ax, width=0.8, alpha=0.9, rot=0)

        # 调整颜色：新一年用原色，老一年用淡色
        # 既然是总量对比图，我们使用主色调 (默认取'电'的颜色，若无则取列表第一个)
        base_theme_color = ENERGY_COLOR_MAP.get(
            "电", list(ENERGY_COLOR_MAP.values())[0]
        )
        years = sorted(pivot_df.columns.tolist())
        max_year = max(years)

        for i, container in enumerate(ax.containers):
            current_year = years[i]
            is_newest = current_year == max_year
            color = (
                base_theme_color
                if is_newest
                else lighten_color(base_theme_color, amount=0.5)
            )
            for bar in container:
                bar.set_facecolor(color)

        plt.title(title, fontsize=30, pad=30, fontweight="bold")
        plt.xlabel("月份", fontsize=24, labelpad=15)
        plt.ylabel(ylabel, fontsize=24, labelpad=15)

        if use_wan_yuan:
            ax.yaxis.set_major_formatter(
                FuncFormatter(lambda value, _: f"{value / 10000:.1f}")
            )
            plt.ylabel(f"{ylabel} (万元)", fontsize=24, labelpad=15)
        else:
            ax.yaxis.set_major_formatter(
                FuncFormatter(lambda value, _: f"{value:,.0f}")
            )

        plt.xticks(fontsize=20)
        plt.yticks(fontsize=20)

        # Legend
        plt.legend(
            title="年份",
            fontsize=20,
            title_fontsize=22,
            loc="upper left",
            bbox_to_anchor=(1.01, 1),
        )

        # Add labels
        for c in ax.containers:
            labels = []
            for v in c:
                height = v.get_height()
                if height > 0:
                    if use_wan_yuan:
                        labels.append(f"{height / 10000:,.1f}")
                    else:
                        labels.append(f"{height:,.0f}")
                else:
                    labels.append("")
            ax.bar_label(
                c,  # type: ignore
                labels=labels,
                padding=3,
                fontsize=16,
                rotation=45,
            )

        plt.tight_layout()
        output_path = os.path.join(output_dir, filename)
        plt.savefig(output_path)
        plt.close()

        logger.info(f"已生成对比分析图: {output_path}")
        print(f"已生成对比分析图: {output_path}")

    except Exception as e:
        logger.error(f"生成对比图表 {filename} 失败: {e}", exc_info=True)


def generate_multi_energy_comparison_bar(
    df: pd.DataFrame,
    output_dir: str,
    value_col: str = "费用(元)",
    title: str = "年度能源费用同比分析",
    ylabel: str = "数值",
    filename: str = "comparison_all_energies.png",
    use_wan_yuan: bool = False,
):
    """
    生成多能源类型对比柱状图。
    X 轴为能源类型，不同年份为并列柱子。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    if df is None or df.empty:
        return

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 1. 数据准备
        plot_df_raw = df[~df["日期区间"].str.contains("合计|总计", na=False)].copy()

        # Pivot: Index=能源类型, Columns=年份
        pivot_df = plot_df_raw.pivot_table(
            index="能源类型",
            columns="年份",
            values=value_col,
            aggfunc="sum",
        )

        if pivot_df.empty:
            logger.warning(f"对比图表 {filename} 无有效数据。")
            return

        # 2. 绘图
        plt.figure(figsize=(20, 12))
        ax = plt.gca()

        # 绘制柱状图
        pivot_df.plot(kind="bar", ax=ax, width=0.8, alpha=0.9, rot=0)

        # 按照用户需求调整颜色：新一年用原色，老一年用淡色
        years = sorted(pivot_df.columns.tolist())
        energy_types = pivot_df.index.tolist()
        max_year = max(years)

        for i, container in enumerate(ax.containers):
            # 每个 container 对应一个年份 (Series)
            current_year = years[i]
            is_newest = current_year == max_year

            for j, bar in enumerate(container):
                # 每个 bar 对应一个能源类型
                etype = energy_types[j]
                base_color = ENERGY_COLOR_MAP.get(etype, "#999999")

                if is_newest:
                    bar.set_facecolor(base_color)
                else:
                    # 老年份使用淡化色
                    bar.set_facecolor(lighten_color(base_color, amount=0.5))

        # 标题中加入月份说明
        months_list = df["日期区间"].unique()
        months_text = ",".join(months_list)
        plt.title(f"{title}\n({months_text})", fontsize=30, pad=35, fontweight="bold")

        plt.xlabel("能源类型", fontsize=24, labelpad=15)

        if use_wan_yuan:
            ax.yaxis.set_major_formatter(
                FuncFormatter(lambda value, _: f"{value / 10000:.1f}")
            )
            plt.ylabel(f"{ylabel} (万元)", fontsize=24, labelpad=15)
        else:
            ax.yaxis.set_major_formatter(
                FuncFormatter(lambda value, _: f"{value:,.0f}")
            )
            plt.ylabel(ylabel, fontsize=24, labelpad=15)

        plt.xticks(fontsize=22)
        plt.yticks(fontsize=20)
        plt.legend(title="年份", fontsize=20, title_fontsize=22)

        # 添加标签
        for c in ax.containers:
            labels = []
            for v in c:
                height = v.get_height()
                if height > 0:
                    label_text = (
                        f"{height / 10000:,.1f}" if use_wan_yuan else f"{height:,.0f}"
                    )
                    labels.append(label_text)
                else:
                    labels.append("")
            ax.bar_label(c, labels=labels, padding=3, fontsize=18, fontweight="bold")

        plt.grid(axis="y", linestyle="--", alpha=0.7)
        plt.tight_layout()

        output_path = os.path.join(output_dir, filename)
        plt.savefig(output_path)
        plt.close()
        logger.info(f"已生成多能源汇总对比图: {output_path}")

    except Exception as e:
        logger.error(f"生成多能源对比图失败: {e}", exc_info=True)


def generate_comparison_pie_charts(
    df: pd.DataFrame,
    output_dir: str,
    value_col: str = "费用(元)",
    title_prefix: str = "能源费用分布同比",
    filename: str = "comparison_pie_distribution.png",
):
    """
    在一个图片中生成不同年份的对比饼图（并排显示）。
    底部注明总量。
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    if df is None or df.empty:
        return

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 1. 数据准备
        plot_df_raw = df[~df["日期区间"].str.contains("合计|总计", na=False)].copy()
        years = sorted(plot_df_raw["年份"].unique())

        if not years:
            return

        # 获取所有能源类型用于统一图例
        all_labels = sorted(plot_df_raw["能源类型"].unique())

        # 2. 绘图设置
        fig, axes = plt.subplots(1, len(years), figsize=(12 * len(years), 10))
        if len(years) == 1:
            axes = [axes]

        unit = "元"
        if value_col == "标准煤(吨标准煤)":
            unit = "吨标煤"

        for i, year in enumerate(years):
            ax = axes[i]
            year_df = plot_df_raw[plot_df_raw["年份"] == year]

            # 汇总该年份各能源类型数据
            year_data = year_df.groupby("能源类型")[value_col].sum()
            year_data = year_data[year_data > 0]

            labels = year_data.index.tolist()
            values = year_data.values.tolist()
            total_val = sum(values)

            if not values:
                ax.text(0.5, 0.5, f"{year}年 无数据", ha="center")
                ax.axis("off")
                continue

            # 使用全局颜色映射
            colors = [ENERGY_COLOR_MAP.get(l, "#999999") for l in labels]

            wedges, texts, autotexts = ax.pie(
                values,
                labels=None,
                autopct=lambda p: f"{p:.1f}%" if p >= 2 else "",
                startangle=140,
                colors=colors,
                pctdistance=0.8,
                textprops={"fontsize": 16, "fontweight": "bold"},
            )

            ax.set_title(f"{year}年", fontsize=26, fontweight="bold", pad=10)

            # 底部标注总量
            val_str = (
                f"{total_val/10000:.2f} 万元"
                if "费用" in value_col and total_val > 10000
                else f"{total_val:,.2f} {unit}"
            )
            ax.text(
                0.5,
                -0.1,
                f"总量: {val_str}",
                transform=ax.transAxes,
                ha="center",
                fontsize=22,
                fontweight="bold",
            )

        # 3. 统一图例
        legend_elements = [
            Patch(facecolor=ENERGY_COLOR_MAP.get(l, "#999999"), label=l)
            for l in all_labels
        ]
        fig.legend(
            handles=legend_elements,
            title="能源类型",
            loc="center right",
            fontsize=16,
            title_fontsize=18,
        )

        # 标题中加入月份说明
        months_text = ",".join(df["日期区间"].unique())
        fig.suptitle(
            f"{title_prefix} ({months_text})", fontsize=30, fontweight="bold", y=0.98
        )

        plt.subplots_adjust(right=0.85, bottom=0.15, top=0.85)

        output_path = os.path.join(output_dir, filename)
        plt.savefig(output_path)
        plt.close()
        logger.info(f"已生成同比饼图: {output_path}")

    except Exception as e:
        logger.error(f"生成同比饼图失败: {e}", exc_info=True)


if __name__ == "__main__":
    generate_pie_charts()
    generate_cost_bar_chart()
    generate_grouped_bar_chart()
