# !/usr/bin/python
# -*- coding:utf-8 -*-
"""
████─█──█─████─███─█──█─███─███
█──█─█──█─█──█─█───██─█──█──█──
████─████─█──█─███─█─██──█──███
█────█──█─█──█─█───█──█──█──█──
█────█──█─████─███─█──█─███─███
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
@time:17:41
@month:十二月
@email:thisluckyboy@126.com
"""
# test_database.py
import unittest
from functools import partial
from wei_office_simptool.wei_office_simptool.utils import OpenExcel,eExcel


class TestDatabase(unittest.TestCase):
    def test_connection(self):
        # 在这里编写你的测试代码
        with OpenExcel(r"D:\基础文件夹\Desktop\日常SQL\1.xlsx",
                               r"D:\基础文件夹\Desktop\日常SQL\1.xlsx").my_open() as wb:
            fastwriteWithParameters = partial(wb.fast_write, wb=wb)
            fastwriteWithParameters('sheet1', ((111,),), 18, 3)
            wb.create_new_sheet("sss1")
            fastwriteWithParameters('sss1', ((111,),), 18, 3)

    # 为其他功能添加更多测试用例

if __name__ == '__main__':
    unittest.main()
