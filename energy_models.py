import pandas as pd
import os
import logging


class EnergySheet:
    """
    Represents a single sheet in an energy data Excel file.
    Handles loading, cleaning, and summarizing the data.
    """

    # 定义需要统计的目标能源类型
    TARGET_TYPES = ["电", "采暖热表", "生活热水表", "自来水", "中水", "燃气"]
    REQUIRED_COLS = ["能源类型", "实际消耗", "费用(元)"]

    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.file_name = os.path.basename(file_path)
        self.logger = logging.getLogger(__name__)
        self.raw_df = None
        self.processed_df = None
        self.summary_df = None

        # 初始化时自动加载并处理
        self._load_and_process()

    def _load_and_process(self):
        """加载数据并执行清洗逻辑"""
        try:
            # 1. 读取数据
            self.raw_df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)

            # 2. 检查列名
            if not all(col in self.raw_df.columns for col in self.REQUIRED_COLS):
                self.logger.error(
                    f"Sheet {self.sheet_name} in {self.file_name} missing required columns. Found: {self.raw_df.columns.tolist()}"
                )
                return

            # 3. 数据清洗核心逻辑
            df = self.raw_df.copy()

            # 处理合并单元格：向下填充能源类型 (解决第0行是电，第1行NaN也是电的问题)
            df["能源类型"] = df["能源类型"].ffill()

            # 去除空格
            df["能源类型"] = df["能源类型"].astype(str).str.strip()

            # Ensure '表号' is string to avoid Parquet mixed type issues
            if "表号" in df.columns:
                df["表号"] = df["表号"].astype(str)

            self.processed_df = df

            # 4. 生成汇总数据 (内部先计算好，随时取用)
            self._generate_summary()

        except Exception as e:
            self.logger.error(
                f"Failed to process Sheet {self.sheet_name} in {self.file_name}: {e}",
                exc_info=True,
            )

    def _generate_summary(self):
        """生成分组汇总数据"""
        if self.processed_df is None:
            return

        df = self.processed_df
        # 筛选目标类型并分组求和
        grouped = (
            df[df["能源类型"].isin(self.TARGET_TYPES)]
            .groupby("能源类型")[["实际消耗", "费用(元)"]]
            .sum()
            .reset_index()
        )

        # 添加元数据
        grouped["日期区间"] = self.sheet_name
        grouped["来源文件"] = self.file_name

        self.summary_df = grouped

    def get_summary(self):
        """接口：获取汇总数据 (DataFrame)"""
        return self.summary_df

    def get_details(self):
        """接口：获取清洗后的详细数据 (DataFrame)"""
        return self.processed_df

    def get_total_by_type(self, energy_type):
        """
        获取指定能源类型的总实际消耗和总费用
        :param energy_type: 能源类型名称，如 '电', '自来水'
        :return: (total_usage, total_cost)
        """
        if self.summary_df is None:
            return 0.0, 0.0

        # 筛选对应能源类型的数据
        result = self.summary_df[self.summary_df["能源类型"] == energy_type]

        if result.empty:
            return 0.0, 0.0

        # summary_df 已经是分组汇总过的，直接取第一行即可
        return result["实际消耗"].iloc[0], result["费用(元)"].iloc[0]

    def _get_cache_path(self, output_dir, format="parquet"):
        """生成缓存文件路径"""
        base_name = os.path.splitext(self.file_name)[0]
        # Clean sheet name to be filename safe
        safe_sheet_name = "".join(
            [c for c in self.sheet_name if c.isalnum() or c in (" ", ".", "-", "_")]
        ).strip()
        file_name = f"{base_name}_{safe_sheet_name}.{format}"
        return os.path.join(output_dir, file_name)

    def save_data(self, output_dir, format="parquet"):
        """
        保存清洗后的详细数据到本地
        :param output_dir: 输出目录
        :param format: 格式 'parquet' 或 'csv'
        """
        if self.processed_df is None:
            return None

        os.makedirs(output_dir, exist_ok=True)
        output_path = self._get_cache_path(output_dir, format)

        try:
            if format == "parquet":
                self.processed_df.to_parquet(output_path, index=False)
            elif format == "csv":
                self.processed_df.to_csv(output_path, index=False, encoding="utf-8-sig")

            self.logger.info(f"Saved processed data to {output_path}")
            return output_path
        except Exception as e:
            self.logger.error(f"Failed to save data to {output_path}: {e}")
            return None

    def compare_with_cache(self, output_dir, format="parquet"):
        """
        与本地缓存数据进行比较
        :param output_dir: 缓存目录
        :param format: 格式
        :return: 'NEW' (新数据), 'MATCH' (一致), 'MISMATCH' (不一致), 'ERROR' (比较出错)
        """
        if self.processed_df is None:
            return "ERROR"

        cache_path = self._get_cache_path(output_dir, format)
        if not os.path.exists(cache_path):
            return "NEW"

        try:
            if format == "parquet":
                cached_df = pd.read_parquet(cache_path)
            elif format == "csv":
                cached_df = pd.read_csv(cache_path)
            else:
                return "ERROR"

            # 使用 pandas 的测试工具进行比较，忽略数据类型差异（如 int32 vs int64）
            pd.testing.assert_frame_equal(
                self.processed_df, cached_df, check_dtype=False
            )
            return "MATCH"
        except AssertionError:
            return "MISMATCH"
        except Exception as e:
            self.logger.error(f"Error comparing with cache: {e}")
            return "ERROR"
