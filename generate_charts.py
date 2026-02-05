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
    "采暖热费": "#C1E0A5",
    "生活热水热费": "#F2B3B1",
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
        suffix: 匹配的列后缀 (例如 "_费用(元)" 或 "_标准煤(kg标准煤)")。
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
                    log_parts.append(f"{energy_type}: {val:,.2f}{usage_str}")

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
                unit = "kg标准煤"

            # Create detailed legend labels
            legend_labels = [
                f"{label}: {val:,.2f}{unit}" for label, val in zip(labels, values)
            ]

            # Add legend to the right
            plt.legend(
                wedges,
                legend_labels,
                title="分项明细",
                loc="center left",
                bbox_to_anchor=(0.9, 0, 0.5, 1),
                fontsize=30,
                title_fontsize=30,
            )

            # Add total cost at the bottom
            total_label = "总额" if suffix != "_费用(元)" else "总费用"
            plt.figtext(
                0.5,
                0.05,
                f"{total_label}: {total_cost:,.2f} {unit}",
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

        for col in cost_cols:
            val = df[col].sum()
            # 仅包含总额大于 0 的项
            if val > 0 and col != total_col:
                values.append(val)
                energy_type = col.replace(suffix, "")
                labels.append(energy_type)

                # Log physical usage if available
                usage_col = f"{energy_type}_用量"
                usage_str = ""
                if usage_col in df.columns:
                    usage_val = df[usage_col].sum()
                    usage_str = f" (用量: {usage_val})"
                log_parts.append(f"{energy_type}: {val:,.2f}{usage_str}")

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

        unit = "元" if suffix == "_费用(元)" else "kg标准煤"
        legend_labels = [
            f"{label}: {val:,.2f}{unit}" for label, val in zip(labels, values)
        ]
        plt.legend(
            wedges,
            legend_labels,
            title="分项明细",
            loc="center left",
            bbox_to_anchor=(0.9, 0, 0.5, 1),
            fontsize=30,
            title_fontsize=30,
        )

        plt.figtext(
            0.5,
            0.05,
            f"全年总额: {total_cost:,.2f} {unit}",
            ha="center",
            fontsize=26,
            fontweight="bold",
            color="#333333",
        )

        plt.tight_layout(rect=(0, 0.1, 0.85, 0.95))
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
        full_plot_df = df.set_index("日期区间")[cost_cols].rename(columns=rename_map)

        # Remove rows with 0 total cost
        full_plot_df = full_plot_df[full_plot_df.sum(axis=1) > 0]

        if full_plot_df.empty:
            logger.info("无有效的分组柱状图数据。")
            return

        # Define groups
        major_types = ["电", "采暖热费"]

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
        create_chart_for_columns(major_types, "major", "主要能源")

        # 2. Generate Minor Chart (All others)
        minor_types = [c for c in full_plot_df.columns if c not in major_types]
        create_chart_for_columns(minor_types, "minor", "其他能源")

    except Exception as e:
        logger.error(f"生成分组柱状图失败: {e}", exc_info=True)


if __name__ == "__main__":
    generate_pie_charts()
    generate_cost_bar_chart()
    generate_grouped_bar_chart()
