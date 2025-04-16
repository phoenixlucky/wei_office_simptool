# !/usr/bin/python
# -*- coding:utf-8 -*-
"""
       .__                         .__
______ |  |__   ____   ____   ____ |__|__  ___
\____ \|  |  \ /  _ \_/ __ \ /    \|  \  \/  /
|  |_> >   Y  (  <_> )  ___/|   |  \  |>    <
|   __/|___|  /\____/ \___  >___|  /__/__/\_ \
|__|        \/            \/     \/         \/


╔╗╔╗╔╗╔═══╗╔══╗╔╗──╔══╗╔══╗╔══╗╔═══╗╔══╗
║║║║║║║╔══╝╚╗╔╝║║──╚╗╔╝║╔╗║║╔╗║╚═╗─║╚╗╔╝
║║║║║║║╚══╗─║║─║║───║║─║╚╝║║║║║─╔╝╔╝─║║─
║║║║║║║╔══╝─║║─║║───║║─║╔╗║║║║║╔╝╔╝──║║─
║╚╝╚╝║║╚══╗╔╝╚╗║╚═╗╔╝╚╗║║║║║╚╝║║─╚═╗╔╝╚╗
╚═╝╚═╝╚═══╝╚══╝╚══╝╚══╝╚╝╚╝╚══╝╚═══╝╚══╝

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

佛祖保佑       永不宕机     永无BUG

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
@project:home
@author:Phoenix,weiliaozi
@file:pywork
@ide:PyCharm
@date:2023/12/3
@time:17:35
@month:十二月
@email:thisluckyboy@126.com
"""
from setuptools import setup, find_packages
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()
setup(
    name='wei_office_simptool',
    version='0.1.3',
    author="Ethan Wilkins",
    author_email="thisluckyboy@126.com",
    description="这是一个简化办公工作的高效工具库，专注于提供常见任务的快捷处理能力。它包含以下核心功能模块：1. 数据库操作：支持快速查询、插入、更新、删除等操作，让开发者仅需1到3行代码即可完成复杂的数据库交互。2. Excel处理：提供灵活的Excel读写、数据筛选、格式化功能，轻松生成或解析复杂表格。3. 邮件发送：集成邮件发送功能，支持多附件、HTML内容等自定义配置，邮件交互更轻松。4. 日期与时间戳格式转换：高效支持各种日期格式与时间戳的快速转换与计算，适配多场景需求。5. 文件操作：一键实现文件的移动、复制、重命名等常见操作，提升文件管理效率。通过整合这些功能，工具库致力于帮助用户在最少的代码量下完成繁琐任务，极大提升办公自动化与开发效率。",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/phoenixlucky/wei_office_simptool",
    packages=find_packages(),
    install_requires=[
        'pathlib',
        'pandas',
        'pymysql',
        'datetime',
        'openpyxl',
        'toml',
        'mysql-connector-python',
        'statsmodels',
        'jieba',
        'wordcloud',
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    license="MIT",
    python_requires='>=3.12',
)
