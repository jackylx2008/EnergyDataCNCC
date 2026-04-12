"""
能源数据处理主程序
================

此脚本负责协调整个能源数据的处理流程。主要功能包括：
1. 读取配置文件 (config.yaml)。
2. 扫描输入目录中的 Excel 文件。
3. 使用 EnergySheet 类处理每个工作表的数据。
4. 管理数据缓存 (Parquet 格式)，避免重复处理并检查数据一致性。
5. 生成汇总 Excel 报表，包含各能源类型的费用统计。

使用方法:
    直接运行此脚本: python process_energy_data.py
"""

import os
import yaml
import re
import pandas as pd
import logging
from core.logging_config import setup_logger
from core.energy_models import EnergySheet


def _extract_month_number(label):
    """Extract month number from labels like '1月'."""
    match = re.search(r"(\d+)", str(label))
    return int(match.group(1)) if match else None


def _has_complete_year_summary(summary_df):
    """
    Return True only when:
    1. Months 1-12 all exist.
    2. Every monthly row has non-zero numeric data.
    """
    if summary_df is None or summary_df.empty or "日期区间" not in summary_df.columns:
        return False

    monthly_df = summary_df.copy()
    monthly_df["_month_num"] = monthly_df["日期区间"].apply(_extract_month_number)
    monthly_df = monthly_df[monthly_df["_month_num"].between(1, 12, inclusive="both")]

    month_set = set(monthly_df["_month_num"].tolist())
    if month_set != set(range(1, 13)):
        return False

    numeric_cols = monthly_df.select_dtypes(include="number").columns.tolist()
    numeric_cols = [col for col in numeric_cols if col != "_month_num"]
    if not numeric_cols:
        return False

    monthly_totals = monthly_df[numeric_cols].sum(axis=1)
    return bool((monthly_totals > 0).all())


def process_ledger_file(file_path):
    """
    处理台账格式的 Excel 文件。
    返回一个包含汇总数据的 list of DataFrames。
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)

    mapping = {
        "电": 2,
        "燃气": 8,  # 天然气-成本
        "自来水": 15,  # 自来水-成本
    }

    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # 寻找 '1月' 所在的行
        start_row = -1
        for i in range(len(df_raw)):
            if str(df_raw.iloc[i, 0]).strip() == "1月":
                start_row = i
                break

        if start_row == -1:
            return None

        for i in range(start_row, start_row + 12):
            if i >= len(df_raw):
                break
            row = df_raw.iloc[i]
            month = str(row[0]).strip()
            if not month or "月" not in month or len(month) > 3:
                continue

            for energy_type, col_idx in mapping.items():
                val = 0.0
                if col_idx < len(row):
                    try:
                        raw_val = row[col_idx]
                        val = float(raw_val)
                        if pd.isna(val) or val < 0:
                            val = 0.0
                    except (ValueError, TypeError):
                        val = 0.0

                # Fix for '自来水' (Tap Water): The Excel file often uses a formula (=O*9.5)
                # which isn't cached, resulting in NaN or 0.
                if energy_type == "自来水" and val == 0:
                    try:
                        # Index 14 is Tap Water Volume
                        vol_val = float(row[14])
                        if not pd.isna(vol_val) and vol_val > 0:
                            val = vol_val * 9.5
                    except (ValueError, TypeError, IndexError):
                        pass

                # Always append, even if 0, to ensure columns appear in summary
                results.append(
                    {
                        "日期区间": month,
                        "能源类型": energy_type,
                        "实际消耗": 0.0,
                        "费用(元)": val,
                        "来源文件": file_name,
                    }
                )
        return pd.DataFrame(results)
    except Exception as e:
        logger.error(f"处理台账文件 {file_name} 失败: {e}")
        return None


def _parse_env_file(env_path="common.env"):
    """Parse a minimal .env-style file into a dict."""
    env_values = {}
    if not os.path.exists(env_path):
        return env_values

    with open(env_path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()

            if (
                len(value) >= 2
                and value[0] == value[-1]
                and value[0] in {'"', "'"}
            ):
                value = value[1:-1]

            env_values[key] = value

    return env_values


def _resolve_env_placeholders(value, env_values):
    """Recursively resolve ${ENV_NAME} placeholders in loaded YAML values."""
    if isinstance(value, dict):
        return {
            key: _resolve_env_placeholders(item, env_values)
            for key, item in value.items()
        }

    if isinstance(value, list):
        return [_resolve_env_placeholders(item, env_values) for item in value]

    if not isinstance(value, str):
        return value

    pattern = r"^\$\{([A-Z0-9_]+)\}$"
    match = re.match(pattern, value)
    if not match:
        return value

    env_key = match.group(1)
    if env_key not in env_values:
        logging.getLogger(__name__).warning(
            "环境变量 %s 未设置，保留原占位符。", env_key
        )
        return value

    raw_env_value = env_values[env_key]
    try:
        return yaml.safe_load(raw_env_value)
    except yaml.YAMLError:
        return raw_env_value


def load_config(config_path="config.yaml"):
    """
    加载 YAML 配置文件。

    Args:
        config_path (str): 配置文件路径，默认为 "config.yaml"。

    Returns:
        dict: 包含配置信息的字典。
    """
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    runtime_config = config.get("runtime", {})
    profile = runtime_config.get("profile", "B25B26")
    env_files = runtime_config.get("env_files", {})
    env_path = env_files.get(profile, "common.b25b26.env")

    env_values = _parse_env_file(env_path)
    if not env_values:
        return config

    return _resolve_env_placeholders(config, env_values)


def save_summary(summary_data, output_dir, group_name, save_excel=True):
    """
    汇总分组数据。

    参数:
        summary_data (list): 包含 DataFrames 的列表。
        output_dir (str): 输出目录。
        group_name (str): 分组名称。
        save_excel (bool): 是否保存为 Excel 文件。默认为 True。

    返回:
        tuple: (output_path, pivot_df) 如果未保存文件，output_path 为 None。
    """
    logger = logging.getLogger(__name__)

    if not summary_data:
        return None, None

    logger.info(f"正在汇总分组数据: {group_name}")
    final_df = pd.concat(summary_data, ignore_index=True)

    # Pivot the table to have Energy Types as headers
    pivot_df = final_df.pivot_table(
        index=["日期区间"],
        columns="能源类型",
        values=["费用(元)", "实际消耗"],
        aggfunc="sum",
        fill_value=0,
    )

    # Swap levels and sort to group by Energy Type (e.g., 电_实际消耗, 电_费用)
    pivot_df.columns = pivot_df.columns.swaplevel(0, 1)
    pivot_df.sort_index(axis=1, level=0, inplace=True)

    # Flatten MultiIndex columns
    pivot_df.columns = [f"{col[0]}_{col[1]}" for col in pivot_df.columns]

    # Reset index to make '日期区间' a normal column
    pivot_df.reset_index(inplace=True)

    # 排序逻辑: 尝试按月份数字排序
    def sort_key(x):
        """提取字符串中的数字，用于按自然顺序（1月、2月...）排序"""
        if isinstance(x, str):
            import re

            match = re.search(r"(\d+)", x)
            if match:
                return int(match.group(1))
        return 999

    pivot_df["_sort_key"] = pivot_df["日期区间"].apply(sort_key)
    pivot_df.sort_values("_sort_key", inplace=True)
    pivot_df.drop(columns=["_sort_key"], inplace=True)

    # 计算标准煤 (吨标准煤)
    config = load_config()
    coal_factors = config.get("coal_conversion", {})

    for energy_type, factor in coal_factors.items():
        usage_col = f"{energy_type}_实际消耗"
        if usage_col in pivot_df.columns:
            coal_col = f"{energy_type}_标准煤(吨标准煤)"
            # 计算并保留两位小数 (除以1000换算为吨)
            pivot_df[coal_col] = (pivot_df[usage_col] * factor / 1000.0).round(2)

    # Reorder columns based on specific order, and exclude those with 0 total cost
    ordered_columns = ["日期区间"]
    target_order = ["电", "采暖用热", "生活热水用热", "自来水", "中水", "燃气"]

    for energy_type in target_order:
        col_usage = f"{energy_type}_实际消耗"
        col_cost = f"{energy_type}_费用(元)"
        col_coal = f"{energy_type}_标准煤(吨标准煤)"

        # Check if this category exists and has any non-zero data
        active = False
        if col_cost in pivot_df.columns and pivot_df[col_cost].sum() > 0:
            active = True
        elif col_usage in pivot_df.columns and pivot_df[col_usage].sum() > 0:
            active = True

        if active:
            if col_usage in pivot_df.columns:
                ordered_columns.append(col_usage)
            if col_cost in pivot_df.columns:
                ordered_columns.append(col_cost)
            if col_coal in pivot_df.columns:
                ordered_columns.append(col_coal)

    remaining_cols = [
        col
        for col in pivot_df.columns
        if col not in ordered_columns and pivot_df[col].sum() > 0
    ]
    ordered_columns.extend(remaining_cols)
    pivot_df = pivot_df[ordered_columns]

    # Calculate total cost and total 吨标准煤
    cost_cols = [col for col in pivot_df.columns if col.endswith("_费用(元)")]
    pivot_df["总费用(元)"] = pivot_df[cost_cols].sum(axis=1)

    coal_cols = [col for col in pivot_df.columns if col.endswith("_标准煤(吨标准煤)")]
    pivot_df["总标准煤(吨标准煤)"] = pivot_df[coal_cols].sum(axis=1)

    has_full_year_data = _has_complete_year_summary(pivot_df)
    if has_full_year_data:
        totals = pivot_df.sum(numeric_only=True)
        totals["日期区间"] = "全年合计"
        pivot_df = pd.concat([pivot_df, pd.DataFrame([totals])], ignore_index=True)
    else:
        logger.warning(
            "分组 %s 未满足完整全年统计条件（存在缺失月份或某月整行数据为0），跳过全年合计统计。",
            group_name,
        )

    # 2. 提取年份以获取经营收入
    year = 2025  # 默认 2025
    try:
        source_file = final_df["来源文件"].iloc[0]
        year_match = re.search(r"(20\d{2})", source_file)
        if year_match:
            year = int(year_match.group(1))
    except Exception:
        pass

    # 3. 处理经营收入和万元营收标准煤耗指标
    revenue_data = config.get("operating_revenue", {}).get(year)
    if revenue_data:
        pivot_df["万元营收标准煤耗(kg/万元营收)"] = 0.0

        if isinstance(revenue_data, dict):
            # 存在月度收入数据
            def get_month_num(m_str):
                m = re.search(r"(\d+)", str(m_str))
                return int(m.group(1)) if m else None

            # 计算各个月份的指标
            for idx, row in pivot_df.iterrows():
                m_num = get_month_num(row["日期区间"])
                if m_num in revenue_data:
                    rev = revenue_data[m_num]
                    if rev > 0:
                        efficiency = round(
                            (row["总标准煤(吨标准煤)"] * 1000.0 / rev), 2
                        )
                        pivot_df.loc[idx, "万元营收标准煤耗(kg/万元营收)"] = efficiency

            # 计算全年的指标
            if has_full_year_data:
                total_rev = sum(revenue_data.values())
                if total_rev > 0:
                    total_coal = pivot_df.loc[
                        pivot_df["日期区间"] == "全年合计", "总标准煤(吨标准煤)"
                    ].values[0]
                    total_efficiency = round((total_coal * 1000.0 / total_rev), 2)
                    pivot_df.loc[
                        pivot_df["日期区间"] == "全年合计",
                        "万元营收标准煤耗(kg/万元营收)",
                    ] = total_efficiency

        elif isinstance(revenue_data, (int, float)) and revenue_data > 0:
            # 仅有年度总收入数据，仅计算全年合计行的指标
            if has_full_year_data:
                total_coal = pivot_df.loc[
                    pivot_df["日期区间"] == "全年合计", "总标准煤(吨标准煤)"
                ].values[0]
                total_efficiency = round((total_coal * 1000.0 / revenue_data), 2)
                pivot_df.loc[
                    pivot_df["日期区间"] == "全年合计",
                    "万元营收标准煤耗(kg/万元营收)",
                ] = total_efficiency

    # 4. 处理建筑面积能耗指标
    total_area_config = config.get("total_area", {})
    if total_area_config and has_full_year_data:
        # 仅针对“全年合计”行计算单位面积指标
        # 如果是字典，则取第一个区域作为主要参考，或者为每个区域创建列
        # 为了保持表格简洁，我们为每个配置的区域增加两列
        if isinstance(total_area_config, dict):
            for area_name, area_val in total_area_config.items():
                if area_val > 0:
                    cost_col = f"{area_name}_单位面积费用(元/㎡)"
                    coal_col = f"{area_name}_单位面积标煤(kg/㎡)"
                    pivot_df[cost_col] = 0.0
                    pivot_df[coal_col] = 0.0

                    # 仅更新全年合计行
                    mask = pivot_df["日期区间"] == "全年合计"
                    if mask.any():
                        total_cost = pivot_df.loc[mask, "总费用(元)"].values[0]
                        total_coal = pivot_df.loc[mask, "总标准煤(吨标准煤)"].values[0]
                        pivot_df.loc[mask, cost_col] = round(total_cost / area_val, 2)
                        pivot_df.loc[mask, coal_col] = round(
                            (total_coal * 1000.0) / area_val, 4
                        )
        elif isinstance(total_area_config, (int, float)) and total_area_config > 0:
            pivot_df["单位面积费用(元/㎡)"] = 0.0
            pivot_df["单位面积标煤(kg/㎡)"] = 0.0

            mask = pivot_df["日期区间"] == "全年合计"
            if mask.any():
                total_cost = pivot_df.loc[mask, "总费用(元)"].values[0]
                total_coal = pivot_df.loc[mask, "总标准煤(吨标准煤)"].values[0]
                pivot_df.loc[mask, "单位面积费用(元/㎡)"] = round(
                    total_cost / total_area_config, 2
                )
                pivot_df.loc[mask, "单位面积标煤(kg/㎡)"] = round(
                    (total_coal * 1000.0) / total_area_config, 4
                )

    output_path = None
    if save_excel:
        output_path = os.path.join(output_dir, f"energy_summary_{group_name}.xlsx")
        pivot_df.to_excel(output_path, index=False)
        logger.info(f"分组 {group_name} 汇总已保存至 {output_path}")

    return output_path, pivot_df


def process_reclaimed_water_file(file_path):
    """
    处理中水用量表。
    1. 月份在第0列 (如 1月, 2月)。
    2. 费用在第2列。
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)
    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        for i, row in df_raw.iterrows():
            month = str(row[0]).strip()
            if "月" not in month or len(month) > 3:
                continue

            try:
                # B列(索引1)为用量，C列(索引2)为费用
                vol = 0.0
                cost = 0.0

                def convert_to_float(val):
                    if pd.isna(val):
                        return None
                    try:
                        return float(val)
                    except (ValueError, TypeError):
                        return None

                raw_vol = convert_to_float(row[1]) if len(row) > 1 else None
                raw_cost = convert_to_float(row[2]) if len(row) > 2 else None

                if raw_vol is not None:
                    vol = raw_vol

                if raw_cost is not None:
                    cost = raw_cost
                else:
                    # 如果费用列为空，按 1.0 元/立方米估算
                    cost = vol * 1.0

                if vol >= 0:
                    results.append(
                        {
                            "日期区间": month,
                            "能源类型": "中水",
                            "实际消耗": vol,
                            "费用(元)": cost,
                            "来源文件": file_name,
                        }
                    )
            except (ValueError, TypeError):
                pass

        return pd.DataFrame(results)
    except Exception as e:
        logger.error(f"处理中水文件 {file_name} 失败: {e}")
        return None


def process_heat_station_file(file_path):
    """
    处理热力站月度费用。
    1. 月份在第0列 (1月, 2月, ..., 12月)。
    2. 生活热水热费在第1列。
    3. 采暖热费在第2列。
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)
    results = []

    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=0)

        for i, row in df_raw.iterrows():
            month = str(row.iloc[0]).strip()
            if "月" not in month or "合计" in month or month == "月份":
                continue

            # 生活热水热费 (Col 1)
            try:
                hw_val = float(row.iloc[1])
                if not pd.isna(hw_val) and hw_val >= 0:
                    results.append(
                        {
                            "日期区间": month,
                            "能源类型": "生活热水用热",
                            "实际消耗": 0.0,
                            "费用(元)": hw_val,
                            "来源文件": file_name,
                        }
                    )
            except (ValueError, TypeError):
                pass

            # 采暖热费 (Col 2)
            try:
                h_val = float(row.iloc[2])
                if not pd.isna(h_val) and h_val >= 0:
                    results.append(
                        {
                            "日期区间": month,
                            "能源类型": "采暖用热",
                            "实际消耗": 0.0,
                            "费用(元)": h_val,
                            "来源文件": file_name,
                        }
                    )
            except (ValueError, TypeError):
                pass

        return pd.DataFrame(results)
    except Exception as e:
        logger.error(f"处理热力站文件 {file_name} 失败: {e}")
        return None


def process_consolidated_file(file_path):
    """
    处理合并后的水电气热用量统计表。
    所有数据在一个 ExcelSheet 中。
    mapping 对应 "商管分公司统计水电气热月度基础数据（扣除租户展商）.xlsx"
    """
    file_name = os.path.basename(file_path)
    logger = logging.getLogger(__name__)

    if not os.path.exists(file_path):
        logger.error(f"文件不存在: {file_path}")
        return None

    # 指向费用列索引 (0-based)
    # 格式: {能源类型: (费用列, 用量列)}
    mapping = {
        "电": (2, 1),
        "燃气": (10, 9),
        # "采暖热费": (13, 11),  # 旧逻辑：热力合计
        "自来水": (19, 17),  # 自来水用量在17列 (R列)，18列是单价
        # "生活热水热费": (21, 20), # 旧逻辑：生活热水（实物量为升/立方米）
        "中水": (24, 22),  # 中水用量在22列 (W列)，23列是单价
        "采暖用热": (27, 25),  # 新增：采暖热量 (吉焦)
        "生活热水用热": (30, 28),  # 新增：生活热水热量 (吉焦)
    }

    results = []
    try:
        # Read without header first to find the starting row
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # 寻找 '1月' 所在的行
        start_row = -1
        for i in range(len(df_raw)):
            val = str(df_raw.iloc[i, 0]).strip()
            # 兼容 "1月" 或 "1 月"
            if val.replace(" ", "") == "1月":
                start_row = i
                break

        if start_row == -1:
            logger.error(f"在文件 {file_name} 中未找到 '1月' 开始的行")
            return None

        # Process up to 12 months
        for i in range(start_row, start_row + 12):
            if i >= len(df_raw):
                break

            row = df_raw.iloc[i]
            month = str(row[0]).strip()
            if "月" not in month:
                continue

            for energy_type, (cost_idx, usage_idx) in mapping.items():
                cost = 0.0
                usage = 0.0

                try:
                    if cost_idx < len(row):
                        cost_val = row[cost_idx]
                        cost = float(cost_val) if not pd.isna(cost_val) else 0.0

                    if usage_idx < len(row):
                        usage_val = row[usage_idx]
                        usage = float(usage_val) if not pd.isna(usage_val) else 0.0
                except (ValueError, TypeError):
                    pass

                results.append(
                    {
                        "日期区间": month,
                        "能源类型": energy_type,
                        "实际消耗": usage,
                        "费用(元)": cost,
                        "来源文件": file_name,
                    }
                )

        return pd.DataFrame(results)

    except Exception as e:
        logger.error(f"处理合并文件 {file_name} 失败: {e}")
        return None


def process_energy_cost_workflow(input_file, output_dir):
    """
    工作流: 从能耗费用角度处理数据
    """
    logger = logging.getLogger(__name__)
    group_name = "能耗费用"
    summary_data = []

    if not os.path.exists(input_file):
        logger.error(f"输入文件不存在: {input_file}")
        return {}

    df = process_consolidated_file(input_file)

    if df is not None:
        summary_data.append(df)
    else:
        logger.error("未提取到任何数据。")

    path, df_summary = save_summary(
        summary_data, output_dir, group_name, save_excel=True
    )
    return {group_name: df_summary} if df_summary is not None else {}


def process_phase2_workflow(input_dir, output_dir):
    """
    工作流 2: 处理国会二期能源结算表
    """
    logger = logging.getLogger(__name__)
    group_name = "国会二期"
    summary_data = []

    # 1. 处理主体结算表 (国会二期主体)
    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".xlsx")
        and f.startswith("国会二期主体能源结算计量表")
        and not f.startswith("~$")
    ]

    for file_name in files:
        file_path = os.path.join(input_dir, file_name)
        logger.info(f"[国会二期流] 正在处理主体文件: {file_name}")

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                logger.info(f"  正在处理工作表: {sheet_name}")
                sheet_obj = EnergySheet(file_path, sheet_name)

                # Check cache
                cache_dir = os.path.join(os.path.dirname(output_dir), "data")
                comparison_result = sheet_obj.compare_with_cache(
                    cache_dir, format="parquet"
                )

                if comparison_result == "NEW":
                    logger.info(f"检测到新工作表: {sheet_name}。正在保存到缓存。")
                    sheet_obj.save_data(cache_dir, format="parquet")
                elif comparison_result == "MATCH":
                    logger.info(f"工作表 {sheet_name} 数据一致性检查通过。")  # OK

                summary = sheet_obj.get_summary()
                if summary is not None:
                    summary_data.append(summary)
        except Exception as e:
            logger.error(f"处理文件 {file_name} 失败: {e}", exc_info=True)

    # 2. 处理国会二期相关的中水文件
    base_dir = (
        os.path.dirname(input_dir)
        if os.path.basename(input_dir) == "主体"
        else input_dir
    )
    files_reclaimed = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            # 仅处理文件名包含 "国会二期" 和 "中水用量表" 的文件
            if (
                "中水用量表" in f
                and "国会二期" in f
                and f.endswith(".xlsx")
                and not f.startswith("~$")
            ):
                files_reclaimed.append(os.path.join(root, f))

    for file_path in files_reclaimed:
        file_name = os.path.basename(file_path)
        logger.info(f"[国会二期流] 正在处理中水文件: {file_name}")
        reclaimed_df = process_reclaimed_water_file(file_path)
        if reclaimed_df is not None:
            summary_data.append(reclaimed_df)

    path, df_summary = save_summary(summary_data, output_dir, group_name)
    return {group_name: df_summary} if df_summary is not None else {}


def process_multi_year_comparison(year_file_map, output_dir):
    """
    处理多年数据的对比分析。
    """
    logger = logging.getLogger(__name__)
    combined_data = []

    for year, file_path in year_file_map.items():
        if not os.path.exists(file_path):
            logger.warning(f"年份 {year} 的文件不存在，跳过: {file_path}")
            continue

        # 使用已有的 process_consolidated_file 处理每个合并台账
        df = process_consolidated_file(file_path)
        if df is not None:
            df["年份"] = year  # 注入年份信息
            combined_data.append(df)

    if not combined_data:
        logger.error("未找到任何可用于对比的有效数据。")
        return None

    # 合并所有年份的数据
    full_df = pd.concat(combined_data, ignore_index=True)

    # 统一计算标准煤消耗量（复用 config 里的折算系数）
    config = load_config()
    coal_factors = config.get("coal_conversion", {})

    def calculate_coal(row):
        energy_type = row["能源类型"]
        factor = coal_factors.get(energy_type, 0)
        return row["实际消耗"] * factor / 1000.0

    full_df["标准煤(吨标准煤)"] = full_df.apply(calculate_coal, axis=1).round(4)

    # 保存对比用的中间明细数据
    comparison_excel = os.path.join(output_dir, "energy_comparison_details.xlsx")
    full_df.to_excel(comparison_excel, index=False)
    logger.info(f"年度对比明细已保存至: {comparison_excel}")

    return full_df


def process_excel_files():
    """
    主入口: 依次执行各类独立工作流
    """
    config = load_config()

    # Setup logging
    log_level_str = config.get("log_level", "INFO")
    log_level = getattr(logging, log_level_str.upper(), logging.INFO)
    log_file = config.get("log_file", "./logs/app.log")

    # 这里 setup_logger 只需调用一次，内部实现如果是单例模式则没问题
    # 如果会重复添加 handler，可能需要注意。这里假设 setup_logger 是安全的。
    logger = setup_logger(log_level=log_level, log_file=log_file)

    output_dir = config["paths"]["output_dir"]
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    input_file = None
    input_dir = None
    if "input_file" in config["paths"]:
        input_file = config["paths"]["input_file"]  # New Input File
        input_dir = os.path.dirname(input_file)
    else:
        input_dir = config["paths"].get("input_dir")

    all_summaries = {}

    # 执行能耗费用工作流
    try:
        if input_file:
            res1 = process_energy_cost_workflow(input_file, output_dir)
            all_summaries.update(res1)
        else:
            # Fallback logic if needed, or just skip
            pass
    except Exception as e:
        logger.error(f"能耗费用工作流执行失败: {e}")

    # 执行 国会二期工作流
    try:
        if input_dir:
            res2 = process_phase2_workflow(input_dir, output_dir)
            all_summaries.update(res2)
    except Exception as e:
        logger.error(f"国会二期工作流执行失败: {e}")

    return all_summaries


if __name__ == "__main__":
    process_excel_files()
