## wei_office_simptool

`wei_office_simptool` 一个用于简化办公工作的工具库，提供了数据库操作、Excel 处理、邮件发送、日期时间戳的格式转换、文件移动等常见功能,实现1到3行代码完成相关处理的快捷操作。

#### 🔌安装与升级

使用以下命令安装 `wei_office_simptool`：

```bash
pip install wei_office_simptool
```

使用以下命令升级 `wei_office_simptool`：

```bash
pip install wei_office_simptool --upgrade
```

#### 🔧功能

<!-- #### 1. Database 类 （可以连接各种数据库） 弃用
用于连接和操作数据库。
```python
from wei_office_simptool import Database

# 示例代码
db = Database(host='your_host', port=3306, user='your_user', password='your_password', db='your_database')
result = db("SELECT * FROM your_table", operation_mode="s")
print(result)
``` -->

#### 1. MySQLDatabase 类
主要用于Mysql数据库的快速连接
```python
from wei_office_simptool import MySQLDatabase
```
##### 📌MySQL 连接配置
```python
mysql_config = {
    'host': 'your_host',
    'user': 'your_user',
    'password': 'your_password',
    'database': 'your_database'
}
```
##### ✏️创建 MySQLDatabase 对象
```python
db = MySQLDatabase(mysql_config)
```
##### 📥插入数据
```python
insert_query = "INSERT INTO your_table (column1, column2) VALUES (%s, %s)"
insert_params = ("value1", "value2")
db.execute_query(insert_query, insert_params)
```
##### 🔍查询数据
```python
select_query = "SELECT * FROM your_table"
results = db.fetch_query(select_query)
for row in results:
    print(row)
```
##### ⌛更新数据
```python
update_query = "UPDATE your_table SET column1 = %s WHERE column2 = %s"
update_params = ("new_value", "value2")
db.execute_query(update_query, update_params)
```
##### 🔪删除数据
```python
delete_query = "DELETE FROM your_table WHERE column1 = %s"
delete_params = ("new_value",)
db.execute_query(delete_query, delete_params)
```
##### 🚪关闭连接
```python
db.close()
```
##### SQLAI智能聊天机器人
```python
from wei_office_simptool import SQLManager

# 示例代码
cfg = {
    'user': 'root',
    'password': '你的密码',
    'host': '127.0.0.1',
    'database': 'mlcorpus'
}
db = SQLManager.MySQLDatabase(cfg)
db.run_ai_chatbot(chat_history_size=5, system_msg="System: You are a helpful AI assistant.")
```

#### 2. Excel 相关类
用于创建/读取/写入 Excel，以及通过 Excel 应用刷新数据连接。

```python
from pathlib import Path
from wei_office_simptool import OpenExcel, ExcelHandler, eExcel, ExcelOperation

# 1) 通过 OpenExcel 打开并保存（自动创建不存在文件）
openfile = str(Path.cwd() / "1.xlsx")
savefile = str(Path.cwd() / "2.xlsx")
with OpenExcel(openfile, savefile).my_open() as wb:
    wb.fast_write('sheet1', [[111, 222], [333, 444]], sr=1, sc=1)

# 2) 使用 ExcelHandler 按区块写入/读取
eh = ExcelHandler(savefile)
eh.excel_write('sheet1', [[555]], start_row=3, start_col=3, end_row=3, end_col=3)
rows = eh.excel_read('sheet1', start_row=1, start_col=1, end_row=3, end_col=3)
print(rows)

# 3) 列出工作表并按关键词过滤
sheets = OpenExcel(openfile).file_show(filter=['sheet', '报表'])
print(sheets)

# 4) 将多工作表拆分为多个文件
ExcelOperation(input_file=savefile, output_folder=str(Path.cwd() / "out")).split_table()
```

#### 2.1 eExcel 类
用于快速创建并写入 Excel（不会依赖 Excel 应用）。
```python
from wei_office_simptool import eExcel

wb = eExcel(file_name=r"D:\Desktop\1.xlsx")
data = [[1, 2], [3, 4]]
wb.fast_write(ws="sheet1", results=data, sr=1, sc=1)
readback = wb.excel_read(start_row=1, start_col=1, end_row=2, end_col=2)
print(readback)
```

#### 2.2 快速创建与空表写入
无需手动创建文件或工作表，支持自动创建并写入。
```python
from wei_office_simptool import eExcel, ExcelHandler

# 使用 eExcel.quick 快速创建（不存在则创建）
wb = eExcel.quick(file_name=r"D:\Desktop\quick.xlsx", default_sheet="sheet1")
wb.fast_write(ws="sheet1", results=[[10, 20], [30, 40]], sr=1, sc=1)

# 使用 ExcelHandler 写入不存在的工作表，自动创建
eh = ExcelHandler(r"D:\Desktop\quick.xlsx")
eh.fast_write("new_sheet", [[99]], start_row=1, start_col=1, xl_book=eh)
```

#### 2.3 快速范围写入说明
fast_write 会根据数据自动计算写入范围：
- 当参数 re=0（默认）时，会根据传入的二维数组自动计算结束行列
- 当参数 re=1 时，使用显式传入的 er/ec（结束行列）
```python
# 自动范围计算（re=0）
wb.fast_write(ws="sheet1", results=[[1, 2], [3, 4]], sr=1, sc=1)

# 显式指定范围（re=1）
wb.fast_write(ws="sheet1", results=[[1, 2], [3, 4]], sr=1, sc=1, er=10, ec=10, re=1)
```

#### 2.4 工作表筛选
file_show 支持传入 None、字符串或字符串列表，按关键词过滤工作表名：
```python
from wei_office_simptool import OpenExcel
openfile = r"D:\Desktop\quick.xlsx"

# 全部工作表
print(OpenExcel(openfile).file_show())

# 单关键词
print(OpenExcel(openfile).file_show(filter="sheet"))

# 多关键词
print(OpenExcel(openfile).file_show(filter=["sheet", "报表"]))
```

#### 2.5 常见流水线示例
从创建到写入、刷新连接、拆分保存的一条龙流程：
```python
from pathlib import Path
from wei_office_simptool import eExcel, OpenExcel, ExcelHandler, ExcelOperation

base = Path.cwd()
f = str(base / "pipeline.xlsx")

# 1) 快速创建并写入
wb = eExcel.quick(f, default_sheet="sheet1")
wb.fast_write("sheet1", [[1, 2], [3, 4]], sr=1, sc=1)

# 2) 使用 ExcelHandler 追加写入（自动创建新工作表）
eh = ExcelHandler(f)
eh.fast_write("sheet2", [[5, 6]], start_row=1, start_col=1, xl_book=eh)

# 3) 通过 Excel 应用刷新并保存（需要本机 Excel）
with OpenExcel(f).open_save_Excel() as appwb:
    appwb.api.RefreshAll()

# 4) 拆分工作表到单文件
ExcelOperation(input_file=f, output_folder=str(base / "out")).split_table()
```

#### 3. eSend 类
用于发送邮件。

```python
from wei_office_simptool import eSend

# 示例代码
email_sender = eSend(sender,receiver,username,password,smtpserver='smtp.126.com')
email_sender.send_email(subject='Your Subject', e_content='Your Email Content', file_paths=['/path/to/file/'], file_names=['attachment.txt'])
```

#### 4. DateFormat 类
用于获取最近的时间处理。

```python
from wei_office_simptool import DateFormat

# 示例代码
#timeclass:1日期 date 2时间戳 timestamp 3时刻 time 4datetime
#获取当日的日期字符串
x=DateFormat(interval_day=0,timeclass='date').get_timeparameter(Format="%Y-%m-%d")
print(x)

# 格式化df的表的列属性
df = DateFormat(interval_day=0,timeclass='date').datetime_standar(df, '日期')
```

#### 5. FileManagement 类
用于文件移动并且重命名。
```python
#latest_folder2 当前目录
#destination_directory 目标目录
#target_files2 文件名
#add_prefix 重命名去除数字
#file_type 文件类型
FileManagement().copy_files(latest_folder2, destination_directory, target_files2, rename=True,file_type="xls")
#寻找最新文件夹
latest_folder = FileManagement().find_latest_folder(base_directory)
```

#### 6. StringBaba 类
用于清洗字符串。
```python
from wei_office_simptool import StringBaba

str="""
萝卜
白菜
"""
formatted_str =StringBaba(str1).format_string_sql()
```

#### 7. TextAnalysis 类
用于进行词频分析。
```python
from wei_office_simptool import TextAnalysis
# 示例用法
data = {
    'Category': ['A', 'A', 'B', 'D', 'C'],
    'Text': [
        '我爱自然语言处理',
        '自然语言处理很有趣',
        '机器学习是一门很有前途的学科',
        '我对机器学习很感兴趣',
        '数据科学包含很多有趣的内容'
    ]
}

df = pd.DataFrame(data)

ta = TextAnalysis(df)
result = ta.get_word_freq(group_col='Category', text_col='Text', agg_func=' '.join)

word_freqs = result['word_freq'].tolist()
titles = result['Category'].tolist()

ta.plot_wordclouds(word_freqs, titles)
```
#### 8. ChatBot类 
0.0.29新增，用于连接Ollama的AI接口

```python
from wei_office_simptool import ChatBot

bot = ChatBot(api_url='http://localhost:11434/api/chat')

print("开始聊天（输入 'exit' 退出，输入 'new' 新建聊天）")
while True:
    user_input = input("你: ")
    if user_input.lower() == 'exit':
        break
    elif user_input.lower() == 'new':
        bot.start_new_chat()
        continue

    # 默认使用流式响应，可以根据需要选择非流式响应
    bot.send_message(user_input, stream=True)

print("聊天结束。")
```

## 9 DailyEmailReport 类
用于发送每日报告邮件，支持HTML和纯文本格式。

```python
from wei_office_simptool import DailyEmailReport

# 初始化 DailyEmailReport 实例
email_reporter = DailyEmailReport(
    email_host='smtp.example.com',
    email_port=465,
    email_username='your_email@example.com',
    email_password='your_password'
)

# 添加收件人
email_reporter.add_receiver('recipient@example.com')

# 发送纯文本邮件
text_content = """
Hello,

Here is your daily report.

[Insert your report content here.]

Regards,
Your Name
"""
email_reporter.send_daily_report("Daily Report", text_content)

# 发送HTML邮件 - 方式1
html_content = """
<html>
  <body>
    <h1>Daily Report</h1>
    <p>Hello,</p>
    <p>Here is your <b>daily report</b>.</p>
    <ul>
      <li>Item 1</li>
      <li>Item 2</li>
    </ul>
    <p>Regards,<br>
    Your Name</p>
  </body>
</html>
"""
email_reporter.send_daily_report("HTML Report", html_content, is_html=True)

# 发送HTML邮件 - 方式2
email_reporter.send_daily_report("HTML Report", html_content=html_content)
```

## 贡献
###### 💡有任何问题或建议，请提出 issue。欢迎贡献代码！

##### Copyright (c) 2024 The Python Packaging Authority
 
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

>`The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.`
 
The software is provided "as is," without any warranty of any kind, either express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement. In no event shall the authors or copyright holders be liable for any claims, damages, or other liabilities, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software.

```本软件以“原样”提供，不附带任何形式的明示或暗示保证，包括但不限于对适销性、特定用途适用性以及不侵权的保证。在任何情况下，作者或版权持有者均不对因使用本软件或与本软件的其他交易相关的任何索赔、损害或其他责任承担责任，无论是合同、侵权或其他原因。```

##### 版权和许可
###### © 2026 Ethan Wilkins

###### 该项目基于 MIT 许可证 分发。
