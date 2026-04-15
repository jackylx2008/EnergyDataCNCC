# 能源数据处理与可视化系统

## 项目简介

本项目用于处理能源结算 Excel 台账，完成数据清洗、费用汇总、标准煤折算、同比分析和图表输出。当前包含两条主工作流：

- `main_energy_cost.py`：单年度能耗费用汇总与图表生成
- `main_energy_comparison.py`：跨年度同口径对比，支持指定月份的万元营收标准煤耗对比，以及截至指定月份累计的单位平米指标对比

## 主要功能

- 自动读取 Excel 台账并处理合并单元格等常见格式问题
- 使用 Parquet 缓存中间结果，避免重复处理
- 输出按日期区间、能源类型汇总的费用与标准煤结果
- 生成费用构成图、同比对比图等可视化图表
- 支持按指定月份输出万元营收标准煤耗 Markdown 表格
- 支持输出“截至指定月份累计单位平米指标对比”

## 项目结构

```text
EnergyDataCNCC/
├── input/                     # 输入目录：原始 Excel 文件
├── output/                    # 输出目录：汇总报表和图表
├── data/                      # 缓存目录：Parquet 数据
├── logs/                      # 日志目录
├── config.yaml                # 静态配置
├── common.<profile>.env       # 环境变量配置（示例命名，已忽略）
├── main_energy_cost.py        # 能耗费用工作流入口
├── main_energy_comparison.py  # 年度同比工作流入口
├── core/                      # 核心业务模块
├── tools/                     # 调试/辅助脚本
└── README.md
```

## 环境依赖

- Python 3.8+
- pandas
- matplotlib
- openpyxl
- pyarrow
- pyyaml

安装依赖：

```bash
pip install pandas matplotlib openpyxl pyarrow pyyaml
```

## 配置说明

`config.yaml` 负责静态结构配置，运行时路径、收入、面积、对比月份等值通过 `.env` 提供。配置示例中的文件名和路径建议使用你自己的占位命名，不要直接照搬任何真实生产文件名。

`config.yaml` 示例：

```yaml
runtime:
  profile: PROFILE_A
  env_files:
    PROFILE_A: common.profile_a.env
    PROFILE_B: common.profile_b.env

paths:
  input_file: ${INPUT_FILE}
  output_dir: ${OUTPUT_DIR}

operating_revenue:
  2025: ${OPERATING_REVENUE_2025}
  2026: ${OPERATING_REVENUE_2026}

total_area:
  PROFILE_A: ${TOTAL_AREA_PROFILE_A}

year_comparison:
  files:
    2025: ${YEAR_COMPARISON_FILE_2025}
  months:
    ${YEAR_COMPARISON_MONTHS}
```

`.env` 示例：

```dotenv
LOG_LEVEL=INFO
LOG_FILE=./logs/app.log
INPUT_FILE=./input/sample_energy_ledger.xlsx
OUTPUT_DIR=./output

TOTAL_AREA_PROFILE_A=418680
YEAR_COMPARISON_FILE_2025=./input/comparison_2025.xlsx
YEAR_COMPARISON_FILE_2026=./input/comparison_2026.xlsx
YEAR_COMPARISON_MONTHS=[1, 2, 3]

OPERATING_REVENUE_2025={"1月": 100.0, "2月": 200.0, "3月": 300.0}
OPERATING_REVENUE_2026={"1月": 120.0, "2月": 220.0, "3月": 320.0}
```

说明：

- `runtime.profile` 决定当前读取哪个 `.env`
- `YEAR_COMPARISON_MONTHS` 控制同比月份范围，常用值为 `[1, 2, 3]`
- `main_energy_comparison.py` 中“单位平米指标对比”是按当前对比月份累计口径输出的；当月份为 `[1, 2, 3]` 时，文案显示为“截至3月累计单位平米指标对比”

## 快速开始

1. 将能源结算 Excel 文件放入输入目录，或在 `.env` 中指定输入文件路径。
2. 在 `config.yaml` 中选择 `runtime.profile`。
3. 在对应的 `.env` 文件中填写本地路径、面积、营收和年度对比配置。
4. 运行单年度费用工作流：

```bash
python main_energy_cost.py
```

5. 运行年度同比工作流：

```bash
python main_energy_comparison.py
```

## 输出说明

单年度费用工作流通常输出：

- 汇总表：`output/energy_summary_能耗费用.xlsx`
- 图表目录：`output/charts_能耗费用/`

年度同比工作流通常输出：

- 对比明细：`output/energy_comparison_details.xlsx`
- 对比图表目录：`output/charts_年度对比/`
- 终端摘要：
  - 指定月份万元营收标准煤耗 Markdown 表格
  - 截至指定月份累计单位平米指标对比

## 调试

- `tools/inspect_excel.py`：快速检查新的 Excel 文件结构和列名
- `tools/consolidate_energy_data.py`：辅助整理和核对能源数据
