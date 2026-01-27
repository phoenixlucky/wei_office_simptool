from pathlib import Path
import pandas as pd
import xlwings as xw
import openpyxl
from openpyxl import load_workbook
from contextlib import contextmanager
from .stringManager import StringBaba


class ExcelHandler:
    def __init__(self, file_name):
        self.file_name = file_name
        self.wb = load_workbook(self.file_name)

    def excel_write(self, sheet_name, results, start_row, start_col, end_row, end_col):
        try:
            sheet = self.wb[sheet_name]
            for i, row in enumerate(range(start_row, end_row + 1)):
                for j, value in enumerate(range(start_col, end_col + 1)):
                    sheet.cell(row=row, column=value, value=results[i][j])
            print("Results have been written!")
            self.wb.save(self.file_name)
        except Exception as e:
            print(e)

    def excel_read(self, sheet_name, start_row, start_col, end_row, end_col):
        try:
            sheet = self.wb[sheet_name]
            values = [
                [sheet.cell(row=row, column=col).value for col in range(start_col, end_col + 1)]
                for row in range(start_row, end_row + 1)
            ]
            print("Results have been read!")
            return values
        except Exception as e:
            print(e)

    def excel_save_as(self, file_name2):
        try:
            self.wb.save(file_name2)
            print("The file has been saved as " + str(file_name2))
        except Exception as e:
            print(e)

    def excel_quit(self):
        try:
            self.wb.close()
        except Exception as e:
            print(e)

    @staticmethod
    def fast_write(sheet_name, results, start_row, start_col, end_row=0, end_col=0, re=0, xl_book=None):
        if re == 0 and results:
            end_row = len(results) + start_row - 1
            end_col = len(results[0]) + start_col - 1
        elif re == 1:
            pass
        xl_book.excel_write(sheet_name, results, start_row=start_row, start_col=start_col, end_row=end_row, end_col=end_col)


class OpenExcel:
    def __init__(self, openfile, savefile=None):
        self.openfile = openfile
        self.savefile = savefile

    @contextmanager
    def my_open(self):
        print(f"Opening Excel file: {self.openfile}")
        wb = eExcel(file_name=self.openfile)
        yield wb
        if self.savefile:
            wb.excel_save_as(self.savefile)
        else:
            wb.excel_save_as(self.openfile)

    @contextmanager
    def open_save_Excel(self):
        app = None
        wb = None
        try:
            app = xw.App(visible=False)
            wb = app.books.open(self.openfile)
        except Exception as e:
            if app:
                app.quit()
            raise e
        try:
            yield wb
        finally:
            try:
                wb.api.RefreshAll()
                wb.save(self.savefile or self.openfile)
            finally:
                if app:
                    app.quit()

    def file_show(self, filter=[]):
        app = xw.App(visible=False)
        wb = app.books.open(self.openfile)
        wbsn = wb.sheet_names
        app.quit()
        if filter or filter == [""]:
            wbsn = StringBaba(wbsn).filter_string_list(filter)
        return wbsn


class ExcelOperation:
    def __init__(self, input_file, output_folder):
        self.input_file = input_file
        self.output_folder = output_folder

    def split_table(self):
        excel_file = pd.ExcelFile(self.input_file)
        out_dir = Path(self.output_folder)
        out_dir.mkdir(parents=True, exist_ok=True)
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(self.input_file, sheet_name=sheet_name)
            output_file = f'{sheet_name}.xlsx'
            df.to_excel(str(out_dir / output_file), index=False)


class eExcel:
    def __init__(self, file_name=None):
        self.file_name = file_name
        if not Path(file_name).exists():
            self.create_new_excel(file_name)
        self.wb = openpyxl.load_workbook(file_name)
        self.ws = self.wb.active

    @staticmethod
    def create_new_excel(file_name):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = 'sheet1'
        wb.save(file_name)

    def create_new_sheet(self, ws):
        self.wb.create_sheet(ws)

    def excel_write(self, ws, results, start_row, start_col, end_row, end_col):
        ws_obj = self.wb[ws]
        for i, row in enumerate(range(start_row, end_row + 1)):
            for j, value in enumerate(range(start_col, end_col + 1)):
                ws_obj.cell(row=row, column=value, value=results[i][j])

    def excel_read(self, start_row, start_col, end_row, end_col):
        valueA = [
            [self.ws.cell(row=row, column=col).value for col in range(start_col, end_col + 1)]
            for row in range(start_row, end_row + 1)
        ]
        return valueA

    def excel_save_as(self, file_name2):
        self.wb.save(file_name2)

    def fast_write(self, ws, results, sr, sc, er=0, ec=0, re=0, wb=None):
        if re == 0 and results:
            er = len(results) + sr - 1
            ec = len(results[0]) + sc - 1
        elif re == 1:
            pass
        target = wb if wb else self
        target.excel_write(ws, results, start_row=sr, start_col=sc, end_row=er, end_col=ec)
