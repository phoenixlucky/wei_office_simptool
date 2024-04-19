# wei_office_simptool

`wei_office_simptool` 一个用于简化办公工作的工具库，提供了数据库操作、Excel 处理、邮件发送、日期时间戳的格式转换等常见功能,实现1到3行代码完成相关处理的快捷操作。

## 安装

使用以下命令安装 `wei_office_simptool`：

```bash
pip install wei_office_simptool
```

## 功能

## 1. Database 类
用于连接和操作 MySQL 数据库。
```bash
from wei_office_simptool import Database

# 示例代码
db = Database(host='your_host', port=3306, user='your_user', password='your_password', db='your_database')
result = db("SELECT * FROM your_table", operation_mode="s")
print(result)
```
## 2. ExcelHandler 类
用于处理 Excel 文件，包括写入和读取。

```bash
from wei_office_simptool import OpenExcel,ExcelHandler

# 示例代码
myExcel = OpenExcel(file_name='your_excel_file.xlsx')
with myExcel:
  myExcel.excel_write(result,sheet_name='Sheet1', results=data, start_row=1, start_col=1, end_row=3, end_col=2)
```
## 3. eSend 类
用于发送邮件。

```bash
from wei_office_simptool import eSend

# 示例代码
email_sender = eSend(sender,receiver,username,password,smtpserver='smtp.126.com')
email_sender.send_email(subject='Your Subject', e_content='Your Email Content', file_paths=['/path/to/file/'], file_names=['attachment.txt'])
```

## 4
用于获取最近的时间处理。

```bash
from wei_office_simptool import eConstant

# 示例代码
#timeclass:1日期 2时间戳 3时刻
interval_time = eConstant(interval_day,timeclass).get_timeparameter()
```

## 贡献
#### 有任何问题或建议，请提出 issue。欢迎贡献代码！

Copyright (c) 2024 The Python Packaging Authority
 
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
 
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
### 版权和许可
## © 2024 Ethan Wilkins

### 该项目基于 MIT 许可证 分发。
