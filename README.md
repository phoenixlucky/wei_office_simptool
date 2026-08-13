<div align="center">

# wei-data-shu 🧩

**面向办公自动化与数据处理的 Python 一站式工具库**
**Domain-Oriented Office Automation & Data Utility Toolkit for Python**

[![Python](https://img.shields.io/badge/Python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13%20%7C%203.14-blue?logo=python&logoColor=white)](https://www.python.org/)
[![PyPI version](https://img.shields.io/pypi/v/wei-data-shu?color=blue)](https://pypi.org/project/wei-data-shu/)
[![License](https://img.shields.io/badge/License-GPLv3-blue.svg)](./LICENSE)
[![Development Status](https://img.shields.io/badge/status-beta-yellow)](https://pypi.org/project/wei-data-shu/)
[![GitHub stars](https://img.shields.io/github/stars/phoenixlucky/wei-data-shu?logo=github)](https://github.com/phoenixlucky/wei-data-shu)

---

覆盖 **数据库 · Excel · 文件 · 文本分析 · 数据分析 · 邮件 · AI 对话 · 通用工具** 八大领域，
领域化分包设计，惰性导入零开销，开箱即用。

Covers **eight domains**: database (MySQL), Excel, files, text analytics, data analysis,
email, AI chat (Ollama) and general utilities. Domain-oriented packages with lazy imports
and zero startup overhead — ready to use out of the box.

</div>

---

## ✨ 特性

- **八大领域一站覆盖** — MySQL 数据库、Excel 电子表格、文件管理、文本分析（词频/词云/趋势预测）、数据分析（通用读取/清洗/可视化）、邮件发送、Ollama AI 对话、通用工具
- **领域化分包设计** — 每个领域一个包，根包仅暴露入口，结构清晰、职责单一
- **惰性导入零开销** — 各领域包按需懒加载，`import wei_data_shu` 不拖慢启动
- **开箱即用** — 统一 `wei_data_shu.<domain>` 导入约定，配合完整示例，5 分钟上手
- **可选依赖按需安装** — 文本分析、Excel App 等重量能力通过 `extras` 安装，核心包轻量

## 📖 目录

- [快速开始](#快速开始)
  - [安装](#安装)
  - [导入方式](#导入方式)
  - [命令行工具](#命令行工具)
  - [5 分钟上手](#5-分钟上手)
- [功能概览](#功能概览)
- [项目结构](#项目结构)
- [用法示例](#用法示例)
  - [使用手册（完整分域示例）](docs/USAGE.md)
- [参与贡献](#参与贡献)
- [许可证](#许可证)

---

## 🚀 快速开始

### 安装

```bash
pip install wei-data-shu
```

> 核心包仅依赖 `toml` 与 `requests`，开箱即用；重量能力按需安装：

```bash
# Excel 读写 / 拆分合并（依赖: pandas, openpyxl）
pip install "wei-data-shu[excel]"

# MySQL 数据库（依赖: mysql-connector-python）
pip install "wei-data-shu[database]"

# 文本分析 / 词云 / 趋势预测 / 数据分析（依赖: jieba, numpy, matplotlib, statsmodels, wordcloud, pandas, openpyxl）
pip install "wei-data-shu[analysis]"

# 需要通过本机 Excel 应用操作工作簿（依赖: xlwings + Microsoft Excel）
pip install "wei-data-shu[excel-client]"
```

升级到最新版本：

```bash
pip install --upgrade wei-data-shu
```

### 导入方式

所有公开 API 统一从 `wei_data_shu.<domain>` 导入，根包 `wei_data_shu` 只暴露领域包入口：

```python
from wei_data_shu.database import MySQLDatabase, MySQLDatabaseError
from wei_data_shu.excel import ExcelManager, OpenExcel, ExcelOperation, quick_excel
from wei_data_shu.files import FileManagement
from wei_data_shu.mail import DailyEmailReport
from wei_data_shu.text import DateFormat, StringBaba, TextAnalysis, TrendPredictor
from wei_data_shu.analysis import DataCleaner, read_csv, plot_line, plot_corr_heatmap
from wei_data_shu.ai import ChatBot
from wei_data_shu.utils import fn_timer, generate_password, search_colors
```

### 命令行工具

安装后可直接在终端使用：

```bash
# 查看帮助
wei-data-shu --help
# 或
python -m wei_data_shu --help
```

```bash
# 颜色检索
wei-data-shu colors                # 列出所有颜色
wei-data-shu colors mint           # 按英文名搜索
wei-data-shu colors 薄荷           # 按中文名搜索
wei-data-shu colors "#5BC49F"      # 按 HEX 搜索

# 密码生成
wei-data-shu password --count 10 --length 13

# 日期计算（默认今天，可回退 N 天）
wei-data-shu date                    # 2026-08-13
wei-data-shu date --days 1 --format "%Y%m%d"

# Excel 工作簿信息（需要 excel extras）
wei-data-shu excel info report.xlsx  # 列出各工作表行数
```

### 5 分钟上手

以下示例**不依赖**数据库、邮件服务或本机 Excel，安装后即可运行。它会完成 4 件事：

- 生成当天报表文件名
- 创建一个 Excel 文件并写入示例数据
- 检索颜色表中的中文颜色信息
- 生成一个不含易混淆字符的安全密码

```python
from pathlib import Path

from wei_data_shu.excel import ExcelManager
from wei_data_shu.text import DateFormat
from wei_data_shu.utils import generate_password, search_colors

# 1. 生成日期字符串
today = DateFormat(interval_day=0, timeclass="date").get_timeparameter(Format="%Y-%m-%d")
report_path = Path(f"demo-report-{today}.xlsx")

# 2. 写入 Excel 报表
rows = [
    ["日期", "渠道", "销售额"],
    [today, "电商", 12580],
    [today, "门店", 9680],
    [today, "分销", 7320],
]

with ExcelManager(str(report_path)) as wb:
    wb.write_sheet("日报", rows, start_row=1, start_col=1)
    summary = wb.read_sheet("日报", 1, 1)

# 3. 颜色检索
mint_colors = search_colors("薄荷")

# 4. 密码生成
temp_password = generate_password(13)

# 输出结果
print("报表文件：", report_path.resolve())
print("首行数据：", summary[0])
print("颜色搜索：", mint_colors[0]["hex"], mint_colors[0]["name"], mint_colors[0]["name_zh"])
print("临时密码：", temp_password)
```

运行后你会得到一个 `demo-report-YYYY-MM-DD.xlsx` 文件，并在终端看到类似输出：

```text
报表文件： D:\path\to\demo-report-2026-03-17.xlsx
首行数据： ['日期', '渠道', '销售额']
颜色搜索： #5BC49F mint green 薄荷绿
临时密码： 8rY#FvQ7mK2$T
```

---

## 🧭 功能概览

| 领域 | 导入路径 | 主要 API | 功能 |
| --- | --- | --- | --- |
| 数据库 | `wei_data_shu.database` | `MySQLDatabase`, `MySQLDatabaseError` | MySQL 连接、查询、插入、更新、删除、存储过程 |
| Excel | `wei_data_shu.excel` | `ExcelManager`, `OpenExcel`, `ExcelOperation`, `quick_excel`, `ExcelHandler` | 读写工作簿、样式、DataFrame、工作表管理、拆分合并、Excel App 操作 |
| 文件 | `wei_data_shu.files` | `FileManagement` | 查找最新文件夹、复制文件、批量重命名、删除 |
| 邮件 | `wei_data_shu.mail` | `DailyEmailReport` | SMTP/SSL 发送纯文本/HTML 邮件、附件 |
| 文本 | `wei_data_shu.text` | `DateFormat`, `StringBaba`, `TextAnalysis`, `TrendPredictor`, `MultipleTrendPredictor`, `textCombing` | 日期格式化、字符串清洗、词频分析、词云、ARIMA 趋势预测、段落重组 |
| 数据分析 | `wei_data_shu.analysis` | `read_csv`, `read_json`, `read_excel`, `read_any`, `DataCleaner`, `plot_line`, `plot_bar`, `plot_hist`, `plot_box`, `plot_scatter`, `plot_pie`, `plot_corr_heatmap` | 通用数据读取、缺失值/重复值/异常值处理、归一化、类别编码、常用图表绘制、相关热力图 |
| AI | `wei_data_shu.ai` | `ChatBot` | 对接 Ollama API，支持流式/非流式对话、聊天记录持久化 |
| 工具 | `wei_data_shu.utils` | `fn_timer`, `generate_password`, `search_colors`, `mav_colors` | 函数计时器、安全密码生成、颜色检索 |
| 文档 | `wei_data_shu.docs` | `FileManagement`, `ExcelHandler`, `OpenExcel`, `ExcelOperation` | 文档工作流（Excel + 文件操作的组合编排） |

---

## 🗂 项目结构

```text
wei_data_shu/
├─ wei_data_shu/            # 核心包
│  ├─ __init__.py           # 根包入口，按需惰性加载各个领域包
│  ├─ __main__.py           # python -m 入口
│  ├─ _api.py               # 统一公开 API 注册表
│  ├─ cli.py                # 命令行接口（colors / password / date / excel）
│  ├─ py.typed              # PEP 561 类型标记（IDE 补全）
│  ├─ ai/                   # AI 能力（ChatBot, Ollama）
│  ├─ analysis/             # 数据分析
│  │  ├─ io.py              #   通用数据读取（CSV/JSON/Excel, read_any）
│  │  ├─ cleaning.py        #   DataCleaner（缺失/重复/异常值, 归一化, 编码）
│  │  ├─ charts.py          #   可视化（折线/柱状/直方/箱线/散点/饼图/热力图）
│  │  └─ _deps.py           #   可选依赖守卫
│  ├─ database/             # 数据库能力（MySQL）
│  ├─ docs/                 # 文档工作流（Excel + 文件处理的组合）
│  ├─ excel/                # Excel 能力
│  │  ├─ manager.py         #   核心: ExcelManager
│  │  ├─ handler.py         #   兼容: ExcelHandler
│  │  ├─ client.py          #   桌面: OpenExcel (xlwings)
│  │  ├─ operations.py      #   高级: ExcelOperation (拆分/合并/CSV)
│  │  ├─ quick.py           #   快捷: quick_excel / read_excel_quick
│  │  └─ _helpers.py        #   内部: 样式/创建/自动范围
│  ├─ files/                # 文件处理（FileManagement）
│  ├─ mail/                 # 邮件发送（DailyEmailReport）
│  ├─ text/                 # 文本处理
│  │  ├─ core.py            #   DateFormat, StringBaba, decrypt
│  │  ├─ analysis.py        #   TextAnalysis (jieba 分词, 词云)
│  │  ├─ forecast.py        #   TrendPredictor, MultipleTrendPredictor
│  │  ├─ combiner.py        #   textCombing (段落重组)
│  │  └─ _deps.py           #   可选依赖守卫
│  └─ utils/                # 通用工具
│     ├─ timing.py          #   fn_timer
│     ├─ passwords.py       #   generate_password
│     └─ colors.py          #   mav_colors, search_colors
├─ tests/                   # 单元测试
├─ examples/                # 可运行示例（quickstart / excel / chatbot）
├─ docs/plans/              # 架构设计文档
├─ pyproject.toml           # 包配置 & 依赖
├─ LICENSE                  # GPL-3.0 许可证
└─ README.md                # 本文件
```

### 设计原则

| 原则 | 说明 |
| --- | --- |
| **惰性导入** | 每个领域包使用 `__getattr__` 按需加载，避免启动时全量导入 |
| **统一入口** | 根包只暴露领域包名称，所有公开 API 通过 `wei_data_shu.<domain>.ClassName` 访问 |
| **结构清晰** | 按领域分包，职责明确；`docs` 包编排跨领域的复合工作流 |
| **可选依赖** | Excel / 数据库 / 文本分析 / Excel App 通过 `[excel]` `[database]` `[analysis]` `[excel-client]` extras 按需安装，核心包仅依赖 `toml`/`requests` |

---

## 💻 用法示例

完整的分域示例（数据库 / Excel / 邮件 / 日期 / 字符串 / 文本分析 / 趋势预测 / 文件管理 / AI 对话 / 通用工具 / 数据分析）请参阅 **[📖 使用手册](docs/USAGE.md)**。

这里仅保留一个不依赖任何第三方服务的最小示例：

```python
from wei_data_shu.text import DateFormat
from wei_data_shu.utils import generate_password, search_colors

print(DateFormat(interval_day=1).get_timeparameter())  # 昨天日期
print(generate_password(13))                           # 安全密码
print(search_colors("薄荷")[0])                        # 颜色检索
```

---
## 🗺 Roadmap（计划表）

以下功能已列入规划、尚未实现，欢迎贡献：

| 类别 | 功能 | 说明 | 优先级 |
| --- | --- | --- | --- |
| 数据接入 | SQLite / PostgreSQL 支持 | 数据库层目前仅 MySQL，计划扩展本地零配置的 SQLite 与常用 PostgreSQL | 中 |
| 数据接入 | HTTP/API 数据抓取封装 | 将 `requests` 封装为"拉取 → 解析 → DataFrame"的一站式接口 | 中 |
| 数据接入 | 数据导出封装 | 一键导出 DataFrame 到 CSV / JSON / Excel（多工作表） | 低 |
| 统计分析 | 描述性统计汇总 | 一键输出均值/分位数/偏度/峰度/缺失比例 | 高 |
| 统计分析 | 相关性分析 API | 独立的 Pearson / Spearman / Kendall 相关系数与显著性 | 高 |
| 统计分析 | 假设检验 | t 检验、卡方检验、ANOVA | 中 |
| 统计分析 | 透视表 / 抽样封装 | pivot table、随机抽样、分层抽样 | 中 |
| 建模 | 回归与分类 | 线性回归、逻辑回归封装 | 中 |
| 建模 | 聚类与降维 | KMeans、PCA | 中 |
| 建模 | 通用模型评估 | 分类/回归指标一键计算与交叉验证 | 低 |
| 交付 | 分析报告自动生成 | HTML / Word 模板化报告，自动嵌入图表与统计结论 | 中 |
| 交付 | 图表插入 Excel | 将 matplotlib 图写入 Excel 工作表，打通 analysis 与 excel 领域 | 中 |
| 工程 | 定时任务调度 | 报表/抓取任务的定时执行配置 | 低 |

---

## 🤝 参与贡献

**English:** We welcome contributions! If you have any questions, suggestions, or improvements, please feel free to:

- [Submit an Issue](https://github.com/phoenixlucky/wei-data-shu/issues) — Report bugs or request features
- [Submit a Pull Request](https://github.com/phoenixlucky/wei-data-shu/pulls) — Contribute code

**中文:** 我们欢迎并感谢您的贡献！如果您有任何问题、建议或改进，请随时：

- [提交 Issue](https://github.com/phoenixlucky/wei-data-shu/issues) — 报告 bug 或提出功能建议
- [提交 Pull Request](https://github.com/phoenixlucky/wei-data-shu/pulls) — 贡献代码

---

## 📄 许可证

**Copyright © 2026 Ethan Wilkins.**

**English:** This project is free software: you can redistribute it and/or modify it under the terms of the [GNU General Public License v3 (GPL-3.0)](https://www.gnu.org/licenses/gpl-3.0.html).

**中文:** 本项目为自由软件，您可以依据 [GNU General Public License v3 (GPL-3.0)](https://www.gnu.org/licenses/gpl-3.0.html) 的条款重新分发或修改。

完整的许可证文本请参见项目根目录的 [LICENSE](./LICENSE) 文件。

---

**免责声明 / Disclaimer:**

**English:** This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

**中文:** 本程序按"原样"分发，不附带任何明示或暗示的担保。有关详细信息，请参阅 GNU General Public License。
