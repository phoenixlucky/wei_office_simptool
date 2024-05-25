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
@time:17:33
@month:十二月
@email:thisluckyboy@126.com
"""
import mysql.connector
import openpyxl
import pandas as pd
import pymysql
from .timingTool import fn_timer


class Database:
    def __init__(self, host, port, user, password, db):
        self.connection_state = 0
        self.connection = None
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.db = db
        self.charset = 'utf8'

    def connect(self):
        try:
            if self.connection_state == 0:
                self.connection = pymysql.connect(
                    host=self.host,
                    port=self.port,
                    user=self.user,
                    password=self.password,
                    db=self.db,
                    charset=self.charset,
                    cursorclass=pymysql.cursors.DictCursor
                )
                self.connection_state = 1
            else:
                self.connection_state = 1
        except Exception as e:
            raise Exception(f"Connection failed: {e}")

    def close(self):
        try:
            if self.connection:
                self.connection.cursor().close()
                self.connection.close()
                self.connection_state = 0
        except Exception as e:
            raise Exception(f"Error closing connection: {e}")

    def execute_sql(self, sql, fetch_all=True, df=False, purchases=None, operation_mode="s"):
        try:
            with self.connection.cursor() as cursor:
                if purchases:
                    cursor.executemany(sql, purchases)
                elif operation_mode == "c":
                    # Assuming sql is the stored procedure name
                    cursor.callproc(sql)
                    cursor.connection.commit()
                else:
                    cursor.execute(sql)

                if fetch_all:
                    if df:
                        return pd.DataFrame(cursor.fetchall())
                        cursor.connection.commit()
                    else:
                        return cursor.fetchall()
                        cursor.connection.commit()
                else:
                    self.connection.commit()
        except Exception as e:
            self.connection.rollback()
            raise Exception(f"Error executing SQL: {e}")

    def fetchall(self, sql):
        self.connect()
        return self.execute_sql(sql)

    def writesql(self, sql):
        self.connect()
        self.execute_sql(sql, fetch_all=False)

    def writesqlmany(self, sql, purchases):
        self.connect()
        self.execute_sql(sql, fetch_all=False, purchases=purchases)

    def callsql(self, sql):
        self.connect()
        self.execute_sql(sql, fetch_all=False,operation_mode = "c")

    def to_df(self, sql):
        self.connect()
        return pd.DataFrame(self.execute_sql(sql))

    @fn_timer
    def __call__(self, sql, purchases=None, operation_mode="s", i=0):
        if i == 0:
            if operation_mode == "s":
                return self.fetchall(sql)
            elif operation_mode == "w":
                self.writesql(sql)
            elif operation_mode == "c":
                self.callsql(sql)
            elif operation_mode == "wm":
                self.writesqlmany(sql, purchases)
        else:
            return self.to_df(sql)

class MySQLDatabase:
    def __init__(self, config):
        self.config = config
        self.connection = None
        self.connect()

    def connect(self):
        try:
            self.connection = mysql.connector.connect(**self.config)
            print("Connected to MySQL database")
        except mysql.connector.Error as err:
            print(f"Error: {err}")

    def close(self):
        if self.connection:
            self.connection.close()
            print("MySQL connection closed")

    def execute_query(self, query, params=None):
        cursor = self.connection.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            self.connection.commit()
            print("Query executed successfully")
        except mysql.connector.Error as err:
            print(f"Error: {err}")
        finally:
            cursor.close()

    def fetch_query(self, query, params=None,dictionary=False):
        cursor = self.connection.cursor(dictionary=dictionary)
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            result = cursor.fetchall()
            return result
        except mysql.connector.Error as err:
            print(f"Error: {err}")
        finally:
            cursor.close()




