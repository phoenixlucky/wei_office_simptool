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
@time:17:33
@month:十二月
@email:thisluckyboy@126.com
"""
# 导入需要公开的模块或变量
from .SQLManager import *
from .baseColor import *
from .timingTool import *
from .excelManager import *
from .fileManager import *
from .mailManager import *
from .stringManager import *
from .chartsManager import TrendPredictor,MultipleTrendPredictor,TextAnalysis
from .textManager import *
from .ollamaManager import *

# 定义__all__变量
__all__ = ['SQLManager', 'baseColor', 'timingTool', 'excelManager','fileManager', 'mailManager', 'stringManager','chartsManager',
           'textManager','ollamaManager']

#执行初始化代码
print("wei...The Lord is here, the gods shun...")

# 定义包级别的变量和函数
#package_variable = 123

#def package_function():
#    print("This is a package-level function.")