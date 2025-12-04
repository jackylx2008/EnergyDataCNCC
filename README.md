# 能源数据处理与可视化系统 (Energy Data Processing & Visualization System)

## 项目简介

本项目用于自动化处理能源消耗数据（Excel 格式），进行数据清洗、汇总统计，并生成可视化的费用分析图表。主要用于处理包含电、水、气、热等多种能源类型的结算报表。

## 主要功能

* **自动化数据处理**: 批量读取 Excel 文件，自动处理合并单元格（如“能源类型”列）等格式问题。
* **增量更新与缓存**: 使用 Parquet 格式缓存已处理数据，支持数据一致性比对，避免重复计算。
* **多维度汇总**: 自动生成按日期区间和能源类型的费用汇总报表 (`energy_usage_summary.xlsx`)。
* **可视化图表**:
  * **饼图**: 展示单期能源费用分布。
  * **堆叠柱状图**: 展示各期总费用对比及构成。
  * **分组柱状图**: 对比不同时期各项能源费用的变化。
* **中文支持**: 完善的中文日志记录和图表中文显示。

## 项目结构

```text
EnergyDataCNCC/
├── input/                  # 输入目录：存放原始 Excel 文件
├── output/                 # 输出目录：存放汇总报表和图表
│   └── charts/             # 生成的图片文件
├── data/                   # 缓存目录：存放 Parquet 缓存文件
├── logs/                   # 日志目录
├── config.yaml             # 配置文件
├── main.py                 # 主程序入口
├── process_energy_data.py  # 数据处理核心逻辑
├── energy_models.py        # 数据模型与清洗逻辑
├── generate_charts.py      # 图表生成模块
├── logging_config.py       # 日志配置
├── inspect_excel.py        # 调试工具：检查 Excel 结构
└── README.md               # 项目说明文档
```

## 环境依赖

* Python 3.8+
* Pandas
* Matplotlib
* OpenPyXL (用于读取 Excel)
* PyArrow (用于 Parquet 支持)
* PyYAML

安装依赖:

```bash
pip install pandas matplotlib openpyxl pyarrow pyyaml
```

## 快速开始

1. **准备数据**: 将能源结算 Excel 文件放入 `input/` 目录。
2. **配置**: 检查 `config.yaml` 中的路径配置（通常默认即可）。
3. **运行**:

   ```bash
   python main.py
   ```

4. **查看结果**:
   * 汇总表格: `output/energy_usage_summary.xlsx`
   * 统计图表: `output/charts/`

## 详细说明

### 1. 数据处理 (`process_energy_data.py`)

读取 `input` 目录下的所有 `.xlsx` 文件。针对每个 Sheet（代表一个日期区间），使用 `EnergySheet` 类进行清洗（如处理合并单元格填充）。处理后的数据会与 `data` 目录下的缓存进行比对，确保数据一致性。

### 2. 图表生成 (`generate_charts.py`)

读取生成的汇总 Excel，使用 Matplotlib 绘制图表。已配置 `SimHei` 和 `Microsoft YaHei` 字体以支持中文显示。

### 3. 调试 (`inspect_excel.py`)

如果遇到新的 Excel 格式，可以使用此脚本快速查看文件结构和列名，以便调整代码。

