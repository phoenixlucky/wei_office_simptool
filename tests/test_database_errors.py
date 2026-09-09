import sys
import types
import unittest
from unittest.mock import MagicMock, patch

# mysql-connector-python 属于可选的 database extras，本地可能未安装。
# 注入假模块，使测试聚焦于错误处理逻辑本身。
_fake_mysql = types.ModuleType("mysql")
_fake_connector = MagicMock()
_fake_connector.Error = Exception
_fake_mysql.connector = _fake_connector
sys.modules.setdefault("mysql", _fake_mysql)
sys.modules.setdefault("mysql.connector", _fake_connector)

from wei_data_shu.database import MySQLDatabase, MySQLDatabaseError  # noqa: E402


class _FakeConfig:
    def __getitem__(self, key):
        return key

    def get(self, key, default=None):
        return default

    def keys(self):
        return []


class TestMySQLDatabaseErrors(unittest.TestCase):
    def test_connect_defaults_to_pure_connector_without_mutating_config(self):
        fake = MagicMock()
        config = {"host": "localhost"}

        with patch("wei_data_shu.database.mysql.mysql.connector", fake):
            db = MySQLDatabase.__new__(MySQLDatabase)
            db.config = config
            db.connection = None
            db.connect()

        fake.connect.assert_called_once_with(host="localhost", use_pure=True)
        self.assertEqual(config, {"host": "localhost"})

    def test_connect_respects_explicit_use_pure(self):
        fake = MagicMock()

        with patch("wei_data_shu.database.mysql.mysql.connector", fake):
            db = MySQLDatabase.__new__(MySQLDatabase)
            db.config = {"host": "localhost", "use_pure": False}
            db.connection = None
            db.connect()

        fake.connect.assert_called_once_with(host="localhost", use_pure=False)

    def test_connect_failure_raises_error(self):
        fake = MagicMock()
        fake.connect.side_effect = Exception("connection refused")
        with patch("wei_data_shu.database.mysql.mysql.connector", fake):
            with self.assertRaises(MySQLDatabaseError):
                MySQLDatabase(_FakeConfig())

    def test_fetch_query_failure_raises_error(self):
        fake_conn = MagicMock()
        cursor = MagicMock()
        cursor.execute.side_effect = Exception("bad SQL")
        fake_conn.cursor.return_value = cursor

        db = MySQLDatabase.__new__(MySQLDatabase)
        db.config = {}
        db.connection = fake_conn

        with self.assertRaises(MySQLDatabaseError):
            db.fetch_query("SELECT bad")

    def test_fetch_query_empty_result_returns_list(self):
        fake_conn = MagicMock()
        cursor = MagicMock()
        cursor.fetchall.return_value = []
        fake_conn.cursor.return_value = cursor

        db = MySQLDatabase.__new__(MySQLDatabase)
        db.config = {}
        db.connection = fake_conn

        self.assertEqual(db.fetch_query("SELECT 1"), [])

    def test_context_manager_closes_connection(self):
        fake_conn = MagicMock()
        db = MySQLDatabase.__new__(MySQLDatabase)
        db.config = {}
        db.connection = fake_conn

        with db as entered:
            self.assertIs(entered, db)
        fake_conn.close.assert_called_once()

    def test_operation_without_connection_raises(self):
        db = MySQLDatabase.__new__(MySQLDatabase)
        db.config = {}
        db.connection = None
        with self.assertRaises(MySQLDatabaseError):
            db.fetch_query("SELECT 1")


if __name__ == "__main__":
    unittest.main()
