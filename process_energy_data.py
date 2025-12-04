import os
import yaml
import pandas as pd
import logging
from logging_config import setup_logger
from energy_models import EnergySheet


def load_config(config_path="config.yaml"):
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def process_excel_files():
    # Load configuration
    config = load_config()

    # Setup logging
    log_level_str = config.get("log_level", "INFO")
    log_level = getattr(logging, log_level_str.upper(), logging.INFO)
    log_file = config.get("log_file", "./logs/app.log")
    logger = setup_logger(log_level=log_level, log_file=log_file)

    input_dir = config["paths"]["input_dir"]
    output_dir = config["paths"]["output_dir"]

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"Created output directory: {output_dir}")

    summary_data = []

    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not files:
        logger.warning("No Excel files found in input directory.")
        return

    for file_name in files:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"Processing file: {file_name}")

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                logger.info(f"  Processing sheet: {sheet_name}")

                # Use the EnergySheet class
                sheet_obj = EnergySheet(file_path, sheet_name)

                # Compare with cache and save if new
                cache_dir = os.path.join(os.path.dirname(output_dir), "data")
                comparison_result = sheet_obj.compare_with_cache(
                    cache_dir, format="parquet"
                )

                if comparison_result == "NEW":
                    logger.info(f"New sheet detected: {sheet_name}. Saving to cache.")
                    sheet_obj.save_data(cache_dir, format="parquet")
                elif comparison_result == "MATCH":
                    logger.info(
                        f"Data consistency check passed for sheet {sheet_name}."
                    )
                elif comparison_result == "MISMATCH":
                    logger.error(
                        f"Data mismatch for sheet {sheet_name} in {file_name} compared to cached data! Skipping cache update."
                    )
                else:
                    logger.error(f"Error checking cache for sheet {sheet_name}.")

                summary = sheet_obj.get_summary()

                if summary is not None:
                    summary_data.append(summary)

        except Exception as e:
            logger.error(f"Failed to process file {file_name}: {e}", exc_info=True)

    if summary_data:
        final_df = pd.concat(summary_data, ignore_index=True)

        # Pivot the table to have Energy Types as headers
        pivot_df = final_df.pivot_table(
            index=["日期区间"],
            columns="能源类型",
            values=["费用(元)"],
            aggfunc="sum",
            fill_value=0,
        )

        # Swap levels to group by Energy Type (e.g., 电_费用)
        pivot_df.columns = pivot_df.columns.swaplevel(0, 1)
        pivot_df.sort_index(axis=1, level=0, inplace=True)

        # Flatten MultiIndex columns
        pivot_df.columns = [f"{col[0]}_{col[1]}" for col in pivot_df.columns]

        # Reset index to make '日期区间' a normal column
        pivot_df.reset_index(inplace=True)

        # Reorder columns based on specific order
        ordered_columns = ["日期区间"]
        target_order = ["电", "采暖热表", "生活热水表", "自来水", "中水", "燃气"]

        for energy_type in target_order:
            col_cost = f"{energy_type}_费用(元)"
            if col_cost in pivot_df.columns:
                ordered_columns.append(col_cost)

        # Add any remaining columns that might not be in the target list (just in case)
        remaining_cols = [col for col in pivot_df.columns if col not in ordered_columns]
        ordered_columns.extend(remaining_cols)

        pivot_df = pivot_df[ordered_columns]

        # Calculate total cost
        cost_cols = [col for col in pivot_df.columns if col.endswith("_费用(元)")]
        pivot_df["总费用(元)"] = pivot_df[cost_cols].sum(axis=1)

        output_path = os.path.join(output_dir, "energy_usage_summary.xlsx")
        pivot_df.to_excel(output_path, index=False)
        logger.info(f"Summary saved to {output_path}")
        print(f"Processing complete. Summary saved to {output_path}")
    else:
        logger.warning("No data processed.")
        print("No data processed.")


if __name__ == "__main__":
    process_excel_files()
