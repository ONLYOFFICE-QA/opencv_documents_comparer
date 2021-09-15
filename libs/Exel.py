from rich import print

from libs.helper import Helper
from var import *


class Exel(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        self.copy_for_test(list_of_files)
        self.coordinate = []
        self.errors = []
        self.run_compare_exel(list_of_files)

    # def opener_Exel(path_for_open, file_name):
    #     wb = load_workbook(path, use_iterators=True)
    #     sheet = wb.worksheets[0]
    #
    #     row_count = sheet.max_row
    #     column_count = sheet.max_column
    #     pass

    @staticmethod
    def get_exel_metadata(wb):
        statistics_exel = {
            'num_of_sheets': f'{wb.Sheets.Count}',
        }
        for sh in wb.Sheets:
            ws = wb.Worksheets(sh.Name)
            used = ws.UsedRange
            nrows = used.Row + used.Rows.Count - 1
            ncols = used.Column + used.Columns.Count - 1
            statistics_exel[f'{sh.Name}_nrows'] = nrows
            statistics_exel[f'{sh.Name}_ncols'] = ncols
        return statistics_exel
        pass

    @staticmethod
    def opener_exel(path_for_open, file_name):
        from win32com.client import Dispatch

        xl = Dispatch("Excel.Application")
        xl.Visible = False  # otherwise excel is hidden

        wb = xl.Workbooks.Open(f'{path_for_open}{file_name}')
        statistics_exel = Exel.get_exel_metadata(wb)
        print("count of sheets:", wb.Sheets.Count)

        wb.Close()
        xl.Quit()
        return statistics_exel

    def run_compare_exel(self, list_of_files):
        for file_name in list_of_files:
            file_name_from = file_name.replace(f'.{extension_to}', f'.{extension_from}')
            if extension_to == file_name.split('.')[-1]:
                self.copy(path_to_folder_for_test + file_name_from, path_to_temp_in_test + file_name_from)
                statistics_exel = self.opener_exel(path_to_temp_in_test, file_name_from)
