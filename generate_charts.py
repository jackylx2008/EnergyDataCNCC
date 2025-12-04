import pandas as pd
import matplotlib.pyplot as plt
import os
import logging
from logging_config import setup_logger

# Configure Chinese font support for Matplotlib
plt.rcParams["font.sans-serif"] = [
    "SimHei",
    "Microsoft YaHei",
    "Arial Unicode MS",
]  # Windows/Mac compatible
plt.rcParams["axes.unicode_minus"] = False


def generate_pie_charts():
    # Setup logging
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")

    input_file = "./output/energy_usage_summary.xlsx"
    output_dir = "./output/charts"

    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        df = pd.read_excel(input_file)

        # Identify cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]

        if not cost_cols:
            logger.warning("No cost columns found.")
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
                    # Extract energy type from column name (e.g., "电_费用(元)" -> "电")
                    energy_type = col.replace("_费用(元)", "")
                    labels.append(energy_type)

            if not values:
                logger.info(f"No cost data for {date_range}, skipping chart.")
                continue

            total_cost = sum(values)

            # Create Pie Chart
            # Increase figure size significantly to accommodate larger fonts and legend
            plt.figure(figsize=(16, 10))

            # Pie chart
            # We use a legend to avoid label overlap on the chart itself
            wedges, texts, autotexts = plt.pie(
                values,
                autopct="%1.1f%%",
                startangle=140,
                pctdistance=0.75,
                textprops={"fontsize": 18},  # Increased font size for percentages
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
                fontsize=16,
                title_fontsize=18,
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
            plt.tight_layout(rect=[0, 0.1, 0.85, 0.95])

            # Save chart
            # Clean filename
            safe_date_range = "".join(
                [c for c in str(date_range) if c.isalnum() or c in (" ", ".", "-", "_")]
            ).strip()
            output_path = os.path.join(
                output_dir, f"cost_distribution_{safe_date_range}.png"
            )

            plt.savefig(output_path)
            plt.close()

            logger.info(f"Generated chart: {output_path}")
            print(f"Generated chart: {output_path}")

    except Exception as e:
        logger.error(f"Failed to generate charts: {e}", exc_info=True)


def generate_cost_bar_chart():
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    input_file = "./output/energy_usage_summary.xlsx"
    output_dir = "./output/charts"

    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        return

    try:
        df = pd.read_excel(input_file)

        # Filter cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]
        if not cost_cols:
            logger.warning("No cost columns found for bar chart.")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_费用(元)", "") for col in cost_cols}
        plot_df = df.set_index("日期区间")[cost_cols].rename(columns=rename_map)

        # Filter out rows with 0 total cost
        plot_df = plot_df[plot_df.sum(axis=1) > 0]

        if plot_df.empty:
            logger.info("No valid data for bar chart.")
            return

        # Create figure
        plt.figure(figsize=(18, 12))
        ax = plt.gca()

        # Plot stacked bar chart
        plot_df.plot(kind="bar", stacked=True, ax=ax, width=0.6, alpha=0.9)

        # Styling
        plt.title("各区间能源费用对比", fontsize=30, pad=25)
        plt.xlabel("日期区间", fontsize=24, labelpad=15)
        plt.ylabel("费用 (元)", fontsize=24, labelpad=15)
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
                c,
                labels=labels,
                label_type="center",
                fontsize=14,
                color="white",
                fontweight="bold",
            )

        plt.tight_layout()

        output_path = os.path.join(output_dir, "cost_comparison_bar.png")
        plt.savefig(output_path)
        plt.close()

        logger.info(f"Generated bar chart: {output_path}")
        print(f"Generated bar chart: {output_path}")

    except Exception as e:
        logger.error(f"Failed to generate bar chart: {e}", exc_info=True)


def generate_grouped_bar_chart():
    logger = setup_logger(log_level=logging.INFO, log_file="./logs/charts.log")
    input_file = "./output/energy_usage_summary.xlsx"
    output_dir = "./output/charts"

    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        return

    try:
        df = pd.read_excel(input_file)

        # Filter cost columns
        cost_cols = [col for col in df.columns if col.endswith("_费用(元)")]
        if not cost_cols:
            logger.warning("No cost columns found for grouped bar chart.")
            return

        # Prepare data: Date Range as index, Columns as Energy Types
        rename_map = {col: col.replace("_费用(元)", "") for col in cost_cols}
        plot_df = df.set_index("日期区间")[cost_cols].rename(columns=rename_map)

        # Filter out rows with 0 total cost
        plot_df = plot_df[plot_df.sum(axis=1) > 0]

        if plot_df.empty:
            logger.info("No valid data for grouped bar chart.")
            return

        # Create figure
        plt.figure(figsize=(20, 12))
        ax = plt.gca()

        # Plot grouped bar chart (stacked=False is default)
        plot_df.plot(kind="bar", ax=ax, width=0.8, alpha=0.9)

        # Styling
        plt.title("各区间分项能源费用对比", fontsize=30, pad=25)
        plt.xlabel("日期区间", fontsize=24, labelpad=15)
        plt.ylabel("费用 (元)", fontsize=24, labelpad=15)
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

        # Add value labels on top of each bar
        for c in ax.containers:
            labels = []
            for v in c:
                height = v.get_height()
                if height > 0:
                    labels.append(f"{height:,.0f}")
                else:
                    labels.append("")

            ax.bar_label(
                c,
                labels=labels,
                label_type="edge",
                fontsize=12,
                padding=3,
                rotation=0,  # Keep horizontal for readability unless crowded
            )

        plt.tight_layout()

        output_path = os.path.join(output_dir, "cost_grouped_bar.png")
        plt.savefig(output_path)
        plt.close()

        logger.info(f"Generated grouped bar chart: {output_path}")
        print(f"Generated grouped bar chart: {output_path}")

    except Exception as e:
        logger.error(f"Failed to generate grouped bar chart: {e}", exc_info=True)


if __name__ == "__main__":
    generate_pie_charts()
    generate_cost_bar_chart()
    generate_grouped_bar_chart()
