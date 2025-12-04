import pandas as pd
import os
import logging


class EnergySheet:
    """
    能源数据工作表处理类

    代表 Excel 文件中的一个工作表 (Sheet)。
    负责单个工作表的数据加载、清洗、缓存管理和汇总统计。

    Attributes:
        file_path (str): Excel 文件完整路径。
        sheet_name (str): 工作表名称 (通常代表日期区间)。
        file_name (str): Excel 文件名。
        raw_df (pd.DataFrame): 原始读取的数据。
        processed_df (pd.DataFrame): 清洗后的数据。
        summary_df (pd.DataFrame): 分组汇总后的数据。
    """

    # 定义需要统计的目标能源类型
    TARGET_TYPES = ["电", "采暖热表", "生活热水表", "自来水", "中水", "燃气"]
    REQUIRED_COLS = ["能源类型", "实际消耗", "费用(元)"]

    def __init__(self, file_path, sheet_name):
        """
        初始化 EnergySheet 实例并自动执行加载和处理。

        Args:
            file_path (str): Excel 文件路径。
            sheet_name (str): 工作表名称。
        """
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.file_name = os.path.basename(file_path)
        self.logger = logging.getLogger(__name__)
        self.raw_df: pd.DataFrame | None = None
        self.processed_df: pd.DataFrame | None = None
        self.summary_df: pd.DataFrame | None = None

        # 初始化时自动加载并处理
        self._load_and_process()

    def _load_and_process(self):
        """
        加载数据并执行清洗逻辑。

        步骤:
        1. 读取 Excel 工作表。
        2. 验证必需列是否存在。
        3. 处理合并单元格 (向下填充能源类型)。
        4. 清理数据 (去除空格，转换类型)。
        5. 调用 _generate_summary 生成汇总。
        """
        try:
            # 1. 读取数据
            self.raw_df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)  # type: ignore

            if self.raw_df is None:
                return

            # 2. 检查列名
            if not all(col in self.raw_df.columns for col in self.REQUIRED_COLS):
                self.logger.error(
                    f"文件 {self.file_name} 中的工作表 {self.sheet_name} 缺少必需列。 "
                    f"发现: {self.raw_df.columns.tolist()}"
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
                f"处理文件 {self.file_name} 中的工作表 {self.sheet_name} 失败: {e}",
                exc_info=True,
            )

    def _generate_summary(self):
        """
        生成分组汇总数据。

        根据 TARGET_TYPES 筛选数据，并按能源类型对实际消耗和费用进行求和。
        结果存储在 self.summary_df 中。
        """
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
        """
        获取汇总数据。

        Returns:
            pd.DataFrame: 包含按能源类型分组的消耗和费用汇总数据。
        """
        return self.summary_df

    def get_details(self):
        """
        获取清洗后的详细数据。

        Returns:
            pd.DataFrame: 包含所有原始行（已清洗）的详细数据。
        """
        return self.processed_df

    def get_total_by_type(self, energy_type):
        """
        获取指定能源类型的总实际消耗和总费用。

        Args:
            energy_type (str): 能源类型名称，如 '电', '自来水'。

        Returns:
            tuple: (total_usage, total_cost)
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
        """
        生成缓存文件路径。

        Args:
            output_dir (str): 输出目录路径。
            format (str): 文件格式后缀，默认为 "parquet"。

        Returns:
            str: 完整的缓存文件路径。
        """
        base_name = os.path.splitext(self.file_name)[0]
        # Clean sheet name to be filename safe
        safe_sheet_name = "".join(
            [c for c in self.sheet_name if c.isalnum() or c in (" ", ".", "-", "_")]
        ).strip()
        file_name = f"{base_name}_{safe_sheet_name}.{format}"
        return os.path.join(output_dir, file_name)

    def save_data(self, output_dir, format="parquet"):
        """
        保存清洗后的详细数据到本地文件。

        Args:
            output_dir (str): 输出目录。
            format (str): 保存格式，支持 'parquet' 或 'csv'。

        Returns:
            str | None: 成功保存的文件路径，失败则返回 None。
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

            self.logger.info(f"已保存处理后的数据至 {output_path}")
            return output_path
        except Exception as e:
            self.logger.error(f"保存数据至 {output_path} 失败: {e}")
            return None

    def compare_with_cache(self, output_dir, format="parquet"):
        """
        与本地缓存数据进行比较，以检测数据是否发生变化。

        Args:
            output_dir (str): 缓存目录。
            format (str): 缓存文件格式。

        Returns:
            str: 比较结果状态码。
                - 'NEW': 无缓存，视为新数据。
                - 'MATCH': 数据与缓存一致。
                - 'MISMATCH': 数据与缓存不一致。
                - 'ERROR': 比较过程中出错。
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
                self.processed_df, cached_df, check_dtype=False  # type: ignore
            )
            return "MATCH"
        except AssertionError:
            return "MISMATCH"
        except Exception as e:
            self.logger.error(f"与缓存比较时出错: {e}")
            return "ERROR"
