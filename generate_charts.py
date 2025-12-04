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

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter

try:
    import yaml
except ImportError:  # pragma: no cover - PyYAML optional
    yaml = None

from logging_config import setup_logger

# Configure Chinese font support for Matplotlib
plt.rcParams["font.sans-serif"] = [
    "SimHei",
    "Microsoft YaHei",
    "Arial Unicode MS",
]  # Windows/Mac compatible
plt.rcParams["axes.unicode_minus"] = False

DEFAULT_ENERGY_COLOR_MAP = {
    "电": "#A8C8E1",
    "采暖热表": "#C1E0A5",
    "生活热水表": "#F2B3B1",
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


def generate_pie_charts():
    """
    生成费用分布饼图。

    遍历汇总数据中的每一行 (每个日期区间)，为每个区间生成一个饼图，
    显示不同能源类型的费用占比。

    输出:
        在 output/charts 目录下生成 cost_distribution_{日期区间}.png
    """
    # Setup logging
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    input_file = "./output/energy_usage_summary.xlsx"
    output_dir = "./output/charts"

    if not os.path.exists(input_file):
        logger.error(f"未找到输入文件: {input_file}")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        df = pd.read_excel(input_file)

        # Identify cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]

        if not cost_cols:
            logger.warning("未找到费用列。")
            return

        for index, row in df.iterrows():
            date_range = row["日期区间"]

            # Extract data for this row
            values = []
            labels = []

            for col in cost_cols:
                val = row[col]
                if val > 0:
                    values.append(val)
                    # Example: "电_费用(元)" -> "电"
                    energy_type = col.replace("_费用(元)", "")
                    labels.append(energy_type)

            if not values:
                logger.info(f"{date_range} 无费用数据，跳过图表生成。")
                continue

            total_cost = sum(values)

            # Create Pie Chart
            # Large canvas keeps legend readable
            plt.figure(figsize=(16, 10))

            # Pie chart
            # We use a legend to avoid label overlap on the chart itself
            pie_colors = get_color_sequence(labels)
            wedges, texts, autotexts = plt.pie(  # type: ignore
                values,
                colors=pie_colors,
                autopct="%1.1f%%",
                startangle=140,
                pctdistance=0.75,
                textprops={"fontsize": 18},
            )

            plt.title(
                f"能源费用分布 - {date_range}", fontsize=28, pad=20
            )  # Increased title font size
            plt.axis(
                "equal"
            )  # Equal aspect ratio ensures that pie is drawn as a circle.

            # Create detailed legend labels
            legend_labels = [
                f"{label}: {val:,.2f}元" for label, val in zip(labels, values)
            ]

            # Add legend to the right
            plt.legend(
                wedges,
                legend_labels,
                title="分项费用明细",
                loc="center left",
                bbox_to_anchor=(0.9, 0, 0.5, 1),
                fontsize=30,
                title_fontsize=30,
            )

            # Add total cost at the bottom
            plt.figtext(
                0.5,
                0.05,
                f"总费用: {total_cost:,.2f} 元",
                ha="center",
                fontsize=26,
                fontweight="bold",
                color="#333333",
            )

            # Adjust layout to make room for legend and bottom text
            # rect=[left, bottom, right, top]
            plt.tight_layout(rect=(0, 0.1, 0.85, 0.95))

            # Save chart
            # Clean filename
            allowed_chars = {" ", ".", "-", "_"}
            clean_chars = [
                c for c in str(date_range) if c.isalnum() or c in allowed_chars
            ]
            safe_date_range = "".join(clean_chars).strip()
            output_path = os.path.join(
                output_dir, f"cost_distribution_{safe_date_range}.png"
            )

            plt.savefig(output_path)
            plt.close()

            logger.info(f"已生成图表: {output_path}")
            print(f"已生成图表: {output_path}")

    except Exception as e:
        logger.error(f"生成图表失败: {e}", exc_info=True)


def generate_cost_bar_chart():
    """
    生成费用对比堆叠柱状图。

    以日期区间为 X 轴，费用为 Y 轴，展示各区间的总费用。
    不同能源类型的费用在柱状图中堆叠显示，方便比较总费用及构成。

    输出:
        output/charts/cost_comparison_bar.png
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    input_file = "./output/energy_usage_summary.xlsx"
    output_dir = "./output/charts"

    if not os.path.exists(input_file):
        logger.error(f"未找到输入文件: {input_file}")
        return

    try:
        df = pd.read_excel(input_file)

        # Filter cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]
        if not cost_cols:
            logger.warning("未找到用于柱状图的费用列。")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_费用(元)", "") for col in cost_cols}
        plot_df = df.set_index("日期区间")[cost_cols].rename(columns=rename_map)

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
                fontsize=22,
                fontweight="bold",
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
                fontsize=25,
                color="white",
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


def generate_grouped_bar_chart():
    """
    生成分项费用对比分组柱状图。

    风格参考：
    - 单轴：仅展示分项费用（柱状）。
    - 布局：图例在底部，X轴标签旋转45度。
    - 样式：清晰的背景，数据标签横向显示，字体放大。

    输出:
        output/charts/cost_grouped_bar.png
    """
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    input_file = "./output/energy_usage_summary.xlsx"
    output_dir = "./output/charts"

    if not os.path.exists(input_file):
        logger.error(f"未找到输入文件: {input_file}")
        return

    try:
        df = pd.read_excel(input_file)

        # Filter cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]
        if not cost_cols:
            logger.warning("未找到用于分组柱状图的费用列。")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_费用(元)", "") for col in cost_cols}
        plot_df = df.set_index("日期区间")[cost_cols].rename(columns=rename_map)

        # Filter out rows with 0 total cost
        plot_df = plot_df[plot_df.sum(axis=1) > 0]

        if plot_df.empty:
            logger.info("无有效的分组柱状图数据。")
            return

        # Create figure and primary axis
        fig, ax1 = plt.subplots(figsize=(20, 12))

        # Plot grouped bars on ax1 with consistent palette
        group_colors = get_color_sequence(plot_df.columns.tolist())
        plot_df.plot(
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
            "各区间能源费用统计 (分项)", fontsize=30, pad=30, fontweight="bold"
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
        legend_labels = plot_df.columns.tolist()
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
                c,
                labels=labels,
                label_type="edge",
                fontsize=20,  # Font size set to 20
                padding=3,
                rotation=0,  # Horizontal labels
            )

        output_path = os.path.join(output_dir, "cost_grouped_bar.png")
        plt.savefig(output_path)
        plt.close()

        logger.info(f"已生成分组柱状图: {output_path}")
        print(f"已生成分组柱状图: {output_path}")

    except Exception as e:
        logger.error(f"生成分组柱状图失败: {e}", exc_info=True)


if __name__ == "__main__":
    generate_pie_charts()
    generate_cost_bar_chart()
    generate_grouped_bar_chart()
