"""Excel 常用操作示例：读写、工作表管理、拆分合并。

需要: pip install 'wei-data-shu[excel]'
"""

from pathlib import Path

from wei_data_shu.excel import ExcelManager, ExcelOperation


def manager_demo() -> None:
    path = Path("book.xlsx")
    with ExcelManager(path) as manager:
        manager.create_sheet("数据")
        manager.write_sheet("数据", [["a", "b"], [1, 2], [3, 4]], 1, 1, 3, 2)
        print("工作表:", manager.sheet_names)
        print("内容:", manager.read_sheet("数据", 1, 1, 3, 2))


def operation_demo() -> None:
    op = ExcelOperation("book.xlsx", "output")
    op.split_table(["数据"])  # 拆出 数据.xlsx
    op.convert_to_csv("数据")  # 转 book.csv


if __name__ == "__main__":
    manager_demo()
    operation_demo()
