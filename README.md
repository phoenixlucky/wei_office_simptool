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
## 1.1. MySQLDatabase 类
from wei_office_simptool import MySQLDatabase

#### MySQL 连接配置
    mysql_config = {
        'host': 'your_host',
        'user': 'your_user',
        'password': 'your_password',
        'database': 'your_database'
    }

#### 创建 MySQLDatabase 对象
    db = MySQLDatabase(mysql_config)

#### 插入数据
    insert_query = "INSERT INTO your_table (column1, column2) VALUES (%s, %s)"
    insert_params = ("value1", "value2")
    db.execute_query(insert_query, insert_params)

#### 查询数据
    select_query = "SELECT * FROM your_table"
    results = db.fetch_query(select_query)
    for row in results:
        print(row)

#### 更新数据
    update_query = "UPDATE your_table SET column1 = %s WHERE column2 = %s"
    update_params = ("new_value", "value2")
    db.execute_query(update_query, update_params)

#### 删除数据
    delete_query = "DELETE FROM your_table WHERE column1 = %s"
    delete_params = ("new_value",)
    db.execute_query(delete_query, delete_params)

# 关闭连接
    db.close()

## 2. ExcelHandler 类
用于处理 Excel 文件，包括写入和读取。

```bash
from wei_office_simptool import OpenExcel,ExcelHandler

# 示例代码
     home_file = pathlib.Path.cwd()
     openfile = pathlib.Path(home_file) / "1.xlsx"
     savefile = pathlib.Path(home_file) / "2.xlsx"
     with OpenExcel(openfile, savefile).my_open() as ws:
         eExcel.fast_write(ws, results, sr, sc, er=0, ec=0, re=0)
```
## 3. eSend 类
用于发送邮件。

```bash
from wei_office_simptool import eSend

# 示例代码
email_sender = eSend(sender,receiver,username,password,smtpserver='smtp.126.com')
email_sender.send_email(subject='Your Subject', e_content='Your Email Content', file_paths=['/path/to/file/'], file_names=['attachment.txt'])
```

## 4 eConstant 类
用于获取最近的时间处理。

```bash
from wei_office_simptool import eConstant

# 示例代码
#timeclass:1日期 2时间戳 3时刻
#获取当日的日期字符串
x=eConstant(interval_day=0,timeclass=1).get_timeparameter(Format="%Y-%m-%d")
print(x)
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
