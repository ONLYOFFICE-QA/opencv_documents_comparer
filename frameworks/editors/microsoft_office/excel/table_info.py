# -*- coding: utf-8 -*-

from frameworks.StaticData import StaticData
from frameworks.decorators import async_processing
from host_control import FileUtils, HostInfo

from .handlers import ExcelEvents

if HostInfo().os == 'windows':
    from win32com.client import Dispatch


def handler():
    ExcelEvents().handler_for_thread()


@async_processing(target=handler)
class TableInfo:
    def __init__(self, file_path):
        self.table_info = {}
        self.tmp_dir = StaticData.tmp_dir_in_test
        self.tmp_file = FileUtils.make_tmp_file(file_path, self.tmp_dir)
        self.excel = StaticData.excel
        self.app = Dispatch("Excel.Application")
        self.app.Visible = False
        self.workbook = self.__open(self.tmp_file)

    def __del__(self):
        self.workbook.Close(False)
        self.app.Quit()
        FileUtils.run_command(f"taskkill /t /im {self.excel}")
        FileUtils.delete(self.tmp_file, stdout=False)

    def get(self):
        self.table_info = {'sheets_count': f"{self.workbook.Sheets.Count}"}
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

    def sheets_count(self):
        return self.workbook.Sheets.Count

    def __open(self, file_path: str):
        try:
            return self.app.Workbooks.Open(file_path)
        except Exception as e:
            print(f"Exception when open file: {e}")
