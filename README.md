# wei-data-shu 🌟

[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-GPLv3-blue)](./LICENSE)
[![PyPI](https://img.shields.io/badge/pypi-v0.6.1-green)](https://pypi.org/project/wei-data-shu/)
[![GitHub stars](https://img.shields.io/badge/⭐-wei--data--shu-blue?logo=github)](https://github.com/phoenixlucky/wei-data-shu)

> **🧩 Domain-oriented office automation and data utility toolkit**
> **🧩 面向办公自动化和数据处理的 Python 一站式工具库**
>
> 覆盖 **数据库(MySQL) · Excel · 文件处理 · 文本分析(AI词云) · 邮件发送 · AI对话(Ollama) · 通用工具** 七大领域。
> 领域化分包设计，惰性导入零开销，开箱即用。

---

## 目录

- [快速开始](#快速开始)
  - [安装](#安装)
  - [导入方式](#导入方式)
  - [命令行工具](#命令行工具)
  - [5 分钟上手](#5-分钟上手)
- [功能概览](#功能概览)
- [项目结构](#项目结构)
- [用法示例](#用法示例)
  - [1. MySQLDatabase（数据库）](#1-mysqldatabase数据库)
  - [2. Excel（电子表格）](#2-excel电子表格)
  - [3. DailyEmailReport（邮件）](#3-dailyemailreport邮件)
  - [4. DateFormat（日期处理）](#4-dateformat日期处理)
  - [5. StringBaba（字符串处理）](#5-stringbaba字符串处理)
  - [6. TextAnalysis（文本分析）](#6-textanalysis文本分析)
  - [7. TrendPredictor（趋势预测）](#7-trendpredictor趋势预测)
  - [8. FileManagement（文件管理）](#8-filemanagement文件管理)
  - [9. ChatBot（AI 对话）](#9-chatbotai-对话)
  - [10. Utils（通用工具）](#10-utils通用工具)
- [参与贡献](#参与贡献)
- [许可证](#许可证)

---

## 快速开始

### 安装

```bash
pip install wei-data-shu
```

按需安装可选能力：

```bash
# 文本分析 / 词云 / 趋势预测（依赖: jieba, numpy, matplotlib, statsmodels, wordcloud）
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
from wei_data_shu.database import MySQLDatabase
from wei_data_shu.excel import ExcelManager, OpenExcel, ExcelOperation, quick_excel
from wei_data_shu.files import FileManagement
from wei_data_shu.mail import DailyEmailReport
from wei_data_shu.text import DateFormat, StringBaba, TextAnalysis, TrendPredictor
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

## 功能概览

| 领域 | 导入路径 | 主要 API | 功能 |
| --- | --- | --- | --- |
| 数据库 | `wei_data_shu.database` | `MySQLDatabase` | MySQL 连接、查询、插入、更新、删除、AI 聊天扩展 |
| Excel | `wei_data_shu.excel` | `ExcelManager`, `OpenExcel`, `ExcelOperation`, `quick_excel`, `ExcelHandler` | 读写工作簿、样式、DataFrame、工作表管理、拆分合并、Excel App 操作 |
| 文件 | `wei_data_shu.files` | `FileManagement` | 查找最新文件夹、复制文件、批量重命名、删除 |
| 邮件 | `wei_data_shu.mail` | `DailyEmailReport` | SMTP/SSL 发送纯文本/HTML 邮件、附件 |
| 文本 | `wei_data_shu.text` | `DateFormat`, `StringBaba`, `TextAnalysis`, `TrendPredictor`, `MultipleTrendPredictor`, `textCombing` | 日期格式化、字符串清洗、词频分析、词云、ARIMA 趋势预测、段落重组 |
| AI | `wei_data_shu.ai` | `ChatBot` | 对接 Ollama API，支持流式/非流式对话、聊天记录持久化 |
| 工具 | `wei_data_shu.utils` | `fn_timer`, `generate_password`, `search_colors`, `mav_colors` | 函数计时器、安全密码生成、颜色检索 |
| 文档 | `wei_data_shu.docs` | `FileManagement`, `ExcelHandler`, `OpenExcel`, `ExcelOperation` | 文档工作流（Excel + 文件操作的组合编排） |

---

## 项目结构

```text
wei_data_shu/
├─ wei_data_shu/            # 核心包
│  ├─ __init__.py           # 根包入口，按需惰性加载各个领域包
│  ├─ __main__.py           # python -m 入口
│  ├─ _api.py               # 统一公开 API 注册表
│  ├─ cli.py                # 命令行接口（colors / password）
│  ├─ ai/                   # AI 能力（ChatBot, Ollama）
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
├─ docs/plans/              # 架构设计文档
├─ pyproject.toml           # 包配置 & 依赖
├─ LICENSE                  # GPL-3.0 许可证
└─ README.md                # 本文件
```

设计原则：

- **惰性导入**：每个领域包使用 `__getattr__` 按需加载，避免启动时全量导入
- **统一入口**：根包只暴露领域包名称，所有公开 API 通过 `wei_data_shu.<domain>.ClassName` 访问
- **结构清晰**：按领域分包，职责明确；`docs` 包编排跨领域的复合工作流
- **可选依赖**：文本分析 / Excel App 功能通过 extras 按需安装，核心包轻量

---

## 用法示例

### 1. MySQLDatabase（数据库）

#### 基本 CRUD

```python
from wei_data_shu.database import MySQLDatabase

# 数据库配置
config = {
    "host": "127.0.0.1",
    "port": 3306,
    "user": "root",
    "password": "your_password",
    "database": "your_database",
}

db = MySQLDatabase(config)

# 插入
db.execute_query(
    "INSERT INTO users (name, age) VALUES (%s, %s)",
    ("Alice", 25),
)

# 查询
results = db.fetch_query("SELECT * FROM users WHERE age > %s", (20,))
for row in results:
    print(row)

# 更新
db.execute_query(
    "UPDATE users SET age = %s WHERE name = %s",
    (26, "Alice"),
)

# 删除
db.execute_query("DELETE FROM users WHERE name = %s", ("Bob",))

db.close()
```

#### AI 对话扩展

在数据库连接上直接启用在线的 AI 助手，方便自然语言查询数据库：

```python
from wei_data_shu.database import MySQLDatabase

cfg = {
    "user": "root",
    "password": "your_password",
    "host": "127.0.0.1",
    "port": 3306,
    "database": "mlcorpus",
}
db = MySQLDatabase(cfg)
db.run_ai_chatbot(
    chat_history_size=5,
    system_msg="System: You are a helpful AI assistant. 请用中文回答。",
)
```

---

### 2. Excel（电子表格）

Excel 模块提供 4 个层次的能力：

| 类 | 依赖 | 适用场景 |
| --- | --- | --- |
| `ExcelManager` | 无（openpyxl） | 日常读写、样式、DataFrame、工作表管理 |
| `quick_excel` / `read_excel_quick` | 无（openpyxl） | 极简单次写入 / 读取 |
| `ExcelHandler` | 无（openpyxl） | 旧版兼容接口 |
| `ExcelOperation` | 无（openpyxl + pandas） | 拆分多工作表、合并多个文件、转 CSV |
| `OpenExcel` | xlwings + Microsoft Excel | 调用本机 Excel 应用（刷新公式、宏等） |

推荐优先使用 **`ExcelManager`**。

#### 2.1 ExcelManager — 基本读写

```python
from wei_data_shu.excel import ExcelManager

# 方式一：with 语句（自动保存、关闭）
with ExcelManager("data.xlsx") as wb:
    wb.write_sheet("Sheet1", [["Name", "Age"], ["Alice", 25]], start_row=1, start_col=1)
    wb.fast_write("Sheet1", [["Bob", 30]], start_row=3, start_col=1)
    data = wb.read_sheet("Sheet1", 1, 1)
    print(data)   # [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]

# 方式二：手动管理
wb = ExcelManager("data.xlsx")
wb.fast_write("Sheet1", [[1, 2], [3, 4]], 1, 1)
wb.save()
wb.close()
```

#### 2.2 ExcelManager — DataFrame 读写

```python
import pandas as pd
from wei_data_shu.excel import ExcelManager

df = pd.DataFrame({"Name": ["Alice", "Bob", "Charlie"], "Age": [25, 30, 28]})

with ExcelManager("team.xlsx") as wb:
    wb.write_dataframe("Sheet1", df)

with ExcelManager("team.xlsx") as wb:
    df_read = wb.read_dataframe("Sheet1")
    print(df_read)
```

#### 2.3 ExcelManager — 工作表管理

```python
from wei_data_shu.excel import ExcelManager

wb = ExcelManager("workbook.xlsx")

# 创建新工作表
wb.create_sheet("销售数据")

# 获取工作表信息
info = wb.get_sheet_info("Sheet1")
print(f"行数: {info['max_row']}, 列数: {info['max_column']}")

# 复制工作表
wb.copy_sheet("Sheet1", "Sheet1_备份")

# 删除工作表
wb.delete_sheet("旧数据")

# 列出所有工作表
print(wb.sheet_names)

wb.save()
wb.close()
```

#### 2.4 quick_excel / read_excel_quick（极简模式）

```python
from wei_data_shu.excel import quick_excel, read_excel_quick

# 一行写入
wb = quick_excel("quick.xlsx", [["Name", "Age"], ["Alice", 25], ["Bob", 30]])

# 一行读取（返回列表）
data = read_excel_quick("quick.xlsx")
print(data)

# 读取为 DataFrame
df = read_excel_quick("quick.xlsx", as_dataframe=True)
print(df)
```

#### 2.5 ExcelOperation — 拆分、合并、转 CSV

```python
from wei_data_shu.excel import ExcelOperation

op = ExcelOperation("input.xlsx", "./output")

# 将多工作表的工作簿拆分为单个文件（每个工作表一个 .xlsx）
files = op.split_table()
print("拆分文件:", files)

# 合并多个文件为一个工作簿
op.merge_tables(["sales_q1.xlsx", "sales_q2.xlsx"], "sales_上半年.xlsx")

# 转换为 CSV
csv_path = op.convert_to_csv()
print("CSV 文件:", csv_path)
```

#### 2.6 OpenExcel — 本机 Excel 应用操作

需要安装 Microsoft Excel 和 `xlwings`：`pip install wei-data-shu[excel-client]`

```python
from wei_data_shu.excel import OpenExcel

# 方式一：读写后自动保存
with OpenExcel("data.xlsx").my_open() as wb:
    wb.fast_write("Sheet1", [["Name", "Age"], ["Alice", 25]], 1, 1)

# 方式二：刷新公式（如数据透视表）
with OpenExcel("report.xlsx").open_save_Excel() as appwb:
    appwb.api.RefreshAll()

# 方式三：列出工作簿中的工作表
sheets = OpenExcel("data.xlsx").file_show(filter=["sheet", "报表"])
print(sheets)
```

#### 2.7 完整流水线示例

```python
from pathlib import Path
from wei_data_shu.excel import ExcelManager, OpenExcel, ExcelOperation

base = Path.cwd()
filepath = str(base / "pipeline.xlsx")

# 1. 写入数据
with ExcelManager(filepath) as wb:
    wb.fast_write("Sheet1", [["Name", "Age"], ["Alice", 25], ["Bob", 30]], 1, 1)

# 2. 通过 Excel 应用刷新公式
with OpenExcel(filepath).open_save_Excel() as appwb:
    appwb.api.RefreshAll()

# 3. 拆分工作表
op = ExcelOperation(filepath, str(base / "output"))
op.split_table()

# 4. 转 CSV
csv_file = op.convert_to_csv()
```

---

### 3. DailyEmailReport（邮件）

#### 发送纯文本邮件

```python
from wei_data_shu.mail import DailyEmailReport

email_reporter = DailyEmailReport(
    email_host="smtp.example.com",
    email_port=465,
    email_username="your_email@example.com",
    email_password="your_password",
)

email_reporter.add_receiver("recipient@example.com")

email_reporter.send_daily_report(
    "日报",
    "Hello,\n\n这是今日报表。\n\nBest Regards",
)
```

#### 发送 HTML 邮件

```python
html = """
<html>
  <body>
    <h1>日报</h1>
    <table border="1">
      <tr><th>渠道</th><th>销售额</th></tr>
      <tr><td>电商</td><td>12,580</td></tr>
      <tr><td>门店</td><td>9,680</td></tr>
    </table>
  </body>
</html>
"""
email_reporter.send_daily_report("HTML 日报", html_content=html)
```

#### 发送带附件的邮件

```python
email_reporter.set_email_content(
    subject="带附件的报表",
    body="详见附件。",
    file_paths=["./attachments/"],
    file_names=["report.xlsx"],
)
email_reporter.send_email()
```

---

### 4. DateFormat（日期处理）

#### 生成格式化的日期 / 时间字符串

```python
from wei_data_shu.text import DateFormat

# 今天
today = DateFormat(interval_day=0, timeclass="date").get_timeparameter(Format="%Y-%m-%d")
print(today)  # 2026-03-17

# 昨天
yesterday = DateFormat(interval_day=1, timeclass="date").get_timeparameter(Format="%Y-%m-%d")

# 当前时间（时:分）
now = DateFormat(interval_day=0, timeclass="time").get_timeparameter(Format="%H:%M")
print(now)    # 14:30

# 当前时间戳
ts = DateFormat(interval_day=0, timeclass="timestamp").get_timeparameter()
print(ts)     # time.struct_time(...)

# 当前 datetime 对象
dt = DateFormat(interval_day=0, timeclass="datetime").get_timeparameter()
print(dt)     # datetime.datetime(...)
```

#### 标准化 DataFrame 中的日期列

```python
import pandas as pd
from wei_data_shu.text import DateFormat

df = pd.DataFrame({"日期": ["2026-01-01", "2026/01/02", "2026年1月3日"]})
df = DateFormat(interval_day=0, timeclass="date").datetime_standar(df, "日期")
print(df.dtypes)  # datetime64[ns]
```

---

### 5. StringBaba（字符串处理）

#### SQL 格式化

将多行文本拼接为 SQL `IN` 子句可用的格式：

```python
from wei_data_shu.text import StringBaba

text = """
苹果
香蕉
橘子
"""
result = StringBaba(text).format_string_sql()
print(result)  # "苹果","香蕉","橘子"
```

#### 字符串列表过滤

```python
from wei_data_shu.text import StringBaba

items = ["苹果手机", "香蕉牛奶", "橘子汽水", "笔记本"]
filtered = StringBaba(items).filter_string_list(["手机", "汽水"])
print(filtered)  # ['苹果手机', '橘子汽水']
```

---

### 6. TextAnalysis（文本分析）

需要安装可选依赖：`pip install wei-data-shu[analysis]`

#### 词频分析

```python
import pandas as pd
from wei_data_shu.text import TextAnalysis

data = {
    "Category": ["A", "A", "B", "B", "C"],
    "Text": [
        "我爱自然语言处理",
        "自然语言处理很有趣",
        "机器学习是一门很有前途的学科",
        "深度学习改变了人工智能",
        "数据科学包含统计与编程",
    ],
}
df = pd.DataFrame(data)

ta = TextAnalysis(df)
result = ta.get_word_freq(group_col="Category", text_col="Text", agg_func=" ".join)

print(result[["Category", "word_freq"]])
```

#### 词云绘制

```python
word_freqs = result["word_freq"].tolist()
titles = result["Category"].tolist()
ta.plot_wordclouds(word_freqs, titles, save_path="wordclouds.png")
```

---

### 7. TrendPredictor（趋势预测）

需要安装可选依赖：`pip install wei-data-shu[analysis]`

#### 单序列趋势预测

```python
import pandas as pd
from wei_data_shu.text import TrendPredictor

# 准备数据
dates = pd.date_range(start="2026-01-01", periods=100, freq="D")
values = [100 + i * 0.5 + (i % 7) * 3 for i in range(100)]  # 模拟趋势
df = pd.DataFrame({"日期": dates, "平滑均值": values})

# 创建预测器
predictor = TrendPredictor(
    market_trend_df=df,
    date_col="日期",
    smoothed_avg_col="平滑均值",
    steps=7,           # 预测未来 7 期
    order=(5, 1, 0),   # ARIMA 参数
    freq="D",          # 日频
)

# 查看原始数据（带趋势标签）
print(predictor.original_data())

# 获取预测结果
future_df, forecast, str_forecast, future_dates = predictor.forecast_data()
print(future_df)

# 模型评估
metrics = predictor.cross_validate(test_size=0.2)
print(metrics)

# 模型信息
info = predictor.get_model_info()
print(info)
```

#### 多序列趋势预测

```python
from wei_data_shu.text import MultipleTrendPredictor

# 多列数据，每列是一个独立序列
df_multi = pd.DataFrame({
    "电商": [100, 110, 120, 130, 140, 150, 160],
    "门店": [80, 82, 85, 88, 90, 95, 100],
    "分销": [50, 55, 60, 58, 62, 65, 70],
}, index=pd.date_range(start="2026-01-01", periods=7, freq="D"))

predictor = MultipleTrendPredictor(df_multi, steps=3)
predictions = predictor.predict()
print(predictions)
```

---

### 8. FileManagement（文件管理）

#### 查找最新文件夹

```python
from wei_data_shu.files import FileManagement

latest = FileManagement().find_latest_folder("./backups")
if latest:
    print(f"最新文件夹: {latest}")
```

#### 复制文件

```python
from wei_data_shu.files import FileManagement

fm = FileManagement()

# 复制单个文件
fm.copy_file_simple("./source/report.xlsx", "./dest/report.xlsx")

# 批量复制并重命名（提取文件名中的中文作为新文件名）
fm.copy_files(
    src_dir="./source",
    dest_dir="./dest",
    target_files=["data_2026.xls", "summary_2026.xls"],
    rename=True,
    file_type="xls",
)
```

#### 删除文件 / 文件夹

```python
fm.delete_folder_or_file("./temp/old_data.xlsx")
fm.delete_folder_or_file("./temp/archive")  # 递归删除目录
```

#### 创建文件夹

```python
fm.create_new_folder("./output/reports/2026")
```

---

### 9. ChatBot（AI 对话）

通过 Ollama API 接入本地大语言模型，支持流式输出和聊天记录持久化。

```python
from wei_data_shu.ai import ChatBot

bot = ChatBot(
    api_url="http://localhost:11434/api/chat",
    model="llama3.2",
    messages_file="messages.toml",       # 初始系统提示
    history_file="chat_history.toml",    # 聊天记录自动保存
)

# 流式对话
print("开始聊天（输入 'exit' 退出，输入 'new' 新建会话）")
while True:
    user_input = input("你: ")
    if user_input.lower() == "exit":
        break
    if user_input.lower() == "new":
        bot.start_new_chat()
        continue
    bot.send_message(user_input, stream=True)
```

参数说明：

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `api_url` | — | Ollama API 地址，如 `http://localhost:11434/api/chat` |
| `model` | `llama3.2` | 使用的模型名称 |
| `messages_file` | `messages.toml` | 初始消息配置文件（TOML 格式） |
| `history_file` | `chat_history.toml` | 聊天历史自动保存路径 |
| `stream` | `True` | 是否启用流式输出 |

---

### 10. Utils（通用工具）

#### 函数计时器

用装饰器测量函数执行时间：

```python
from wei_data_shu.utils import fn_timer

@fn_timer
def build_report():
    import time
    time.sleep(0.5)
    return "done"

result, elapsed = build_report()
print(f"耗时: {elapsed:.2f} 秒")  # Total time running build_report: 0.50 seconds
```

#### 密码生成

生成不含易混淆字符（`iIl1o0O`）的安全密码，适合临时密码/一次性密码：

```python
from wei_data_shu.utils import generate_password

# 默认长度 13
pwd = generate_password()
print(pwd)  # 8rY#FvQ7mK2$T

# 自定义长度
pwd16 = generate_password(16)
print(pwd16)

# 批量生成（通过 CLI）
# wei-data-shu password --count 5 --length 16
```

#### 颜色检索

内置 50+ 种常用颜色，支持英文名、中文名、HEX 码检索：

```python
from wei_data_shu.utils import search_colors, mav_colors

# 按英文名搜索
results = search_colors("mint")
print(results[0])
# {'index': 2, 'hex': '#5BC49F', 'name': 'mint green', 'name_zh': '薄荷绿'}

# 按中文名搜索
results = search_colors("薄荷")
print(results[0]["hex"])  # #5BC49F

# 按 HEX 搜索
results = search_colors("#5BC49F")
print(results[0]["name_zh"])  # 薄荷绿

# 查看所有颜色
print(len(mav_colors))   # 39
print(mav_colors[:3])    # ['#60ACFC', '#32D3EB', '#5BC49F']
```

通过 CLI 检索：

```bash
wei-data-shu colors
wei-data-shu colors mint
wei-data-shu colors 薄荷
wei-data-shu colors "#5BC49F"
```

输出格式：

```text
 2. #5BC49F | mint green | 薄荷绿
```

---

## 参与贡献

**English:** We welcome contributions! If you have any questions, suggestions, or improvements, please feel free to:

- [Submit an Issue](https://github.com/phoenixlucky/wei-data-shu/issues) — Report bugs or request features
- [Submit a Pull Request](https://github.com/phoenixlucky/wei-data-shu/pulls) — Contribute code

**中文:** 我们欢迎并感谢您的贡献！如果您有任何问题、建议或改进，请随时：

- [提交 Issue](https://github.com/phoenixlucky/wei-data-shu/issues) — 报告 bug 或提出功能建议
- [提交 Pull Request](https://github.com/phoenixlucky/wei-data-shu/pulls) — 贡献代码

---

## 许可证

**Copyright © 2026 Ethan Wilkins.**

**English:** This project is free software: you can redistribute it and/or modify it under the terms of the [GNU General Public License v3 (GPL-3.0)](https://www.gnu.org/licenses/gpl-3.0.html).

**中文:** 本项目为自由软件，您可以依据 [GNU General Public License v3 (GPL-3.0)](https://www.gnu.org/licenses/gpl-3.0.html) 的条款重新分发或修改。

完整的许可证文本请参见项目根目录的 [LICENSE](./LICENSE) 文件。

---

**免责声明 / Disclaimer:**

**English:** This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

**中文:** 本程序按"原样"分发，不附带任何明示或暗示的担保。有关详细信息，请参阅 GNU General Public License。
