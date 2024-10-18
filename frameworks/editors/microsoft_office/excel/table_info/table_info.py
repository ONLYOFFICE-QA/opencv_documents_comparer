# -*- coding: utf-8 -*-
from functools import wraps

from frameworks.StaticData import StaticData
from frameworks.decorators import async_processing
from host_tools import File, HostInfo, Shell

from frameworks.editors.microsoft_office.excel.handlers import ExcelEvents

if HostInfo().os == 'windows':
    from win32com.client import Dispatch


def handler():
    ExcelEvents().handler_for_thread()

def workbook_exists(method):
    @wraps(method)
    def wrapper(self, *args, **kwargs):
        if self.workbook:
            return method(self, *args, **kwargs)

    return wrapper

@async_processing(target=handler)
class TableInfo:
    def __init__(self, file_path):
        self.table_info = {}
        self.tmp_dir = StaticData.tmp_dir_in_test
        self.tmp_file = File.make_tmp(file_path, self.tmp_dir)
        self.excel = StaticData.excel
        self.app = Dispatch("Excel.Application")
        self.app.Visible = False
        self.workbook = self._open_workbook(self.tmp_file)

    def __del__(self):
        self.close_workbook()
        self._kill_excel_process()
        self._cleanup_tmp_file()

    def get(self):
        self.table_info = {'sheets_count': f"{self.sheets_count()}"}
        try:
            sheet_number = 1
            for sh in self.workbook.Sheets:
                used_range = self.workbook.Worksheets(sh.Name).UsedRange
                self.table_info[f'{sheet_number}_nrows'] = used_range.Row + used_range.Rows.Count - 1
                self.table_info[f'{sheet_number}_ncols'] = used_range.Column + used_range.Columns.Count - 1
                self.table_info[f'{sheet_number}_page_name'] = sh.Name
                sheet_number += 1
        except Exception as e:
            print(f"Exception when get table info: {e}")
        finally:
            return self.table_info

    @workbook_exists
    def close_workbook(self):
        self.workbook.Close(False)
        self.app.Quit()

    @workbook_exists
    def sheets_count(self):
        return self.workbook.Sheets.Count

    def _open_workbook(self, file_path: str):
        try:
            return self.app.Workbooks.Open(file_path)
        except Exception as e:
            print(f"Exception when open file: {e}")
            return None

    def _kill_excel_process(self):
        Shell.call(f"taskkill /t /im {self.excel}")

    def _cleanup_tmp_file(self):
        File.delete(self.tmp_file, stdout=False, stderr=False)
