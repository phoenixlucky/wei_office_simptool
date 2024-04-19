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

from wei_office_simptool.utils import Database


class TestDatabase(unittest.TestCase):
    def test_connection(self):
        # 在这里编写你的测试代码
        db = Database(host='localhost', port=3306, user='user', password='password', db='test_db')
        db.connect()
        self.assertTrue(db.connection_state == 1)
        db.close()
        self.assertTrue(db.connection_state == 0)

    # 为其他功能添加更多测试用例

if __name__ == '__main__':
    unittest.main()
