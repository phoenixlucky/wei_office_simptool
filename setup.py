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
    version='0.0.20',
    author="Ethan Wilkins",
    author_email="thisluckyboy@126.com",
    description="一个用于简化办公工作的工具库，提供了数据库操作、Excel 处理、邮件发送、日期时间戳的格式转换、文件移动等常见功能,实现1到3行代码完成相关处理的快捷操作。",  # 包的简述
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
        # Include other dependencies
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.10',
)
