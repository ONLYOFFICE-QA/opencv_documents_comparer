import csv
import io
import subprocess as sb
import traceback

import win32con
import win32gui
from rich import print
from rich.progress import track
from win32com.client import Dispatch

from libs.helper import Helper
from libs.logger import *
from var import *


class Exel(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        # self.copy_for_test(list_of_files)
        self.coordinate = []
        self.errors = []
        self.run_compare_exel(list_of_files)

    @staticmethod
    def get_exel_metadata(wb):
        statistics_exel = {
            'num_of_sheets': f'{wb.Sheets.Count}',
        }

        num_of_sheet = 1
        for sh in wb.Sheets:
            ws = wb.Worksheets(sh.Name)
            used = ws.UsedRange
            nrows = used.Row + used.Rows.Count - 1
            ncols = used.Column + used.Columns.Count - 1
            statistics_exel[f'{num_of_sheet}_page_name'] = sh.Name
            statistics_exel[f'{num_of_sheet}_nrows'] = nrows
            statistics_exel[f'{num_of_sheet}_ncols'] = ncols
            num_of_sheet += 1
        return statistics_exel
        pass

    def get_windows_title(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword':
                # hwnd = win32gui.FindWindow(None, "Telegram (15125)")
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)

                self.errors.clear()
                self.errors.append(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetWindowText(hwnd))

            # elif win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) == 'Word':
            #     errors.clear()
            #     errors.append(win32gui.GetClassName(hwnd))
            #     errors.append(win32gui.GetWindowText(hwnd))

    @staticmethod
    def opener_exel(path_for_open, file_name):
        try:
            xl = Dispatch("Excel.Application")
            xl.Visible = False  # otherwise excel is hidden
            wb = xl.Workbooks.Open(f'{path_for_open}{file_name}', ReadOnly=True)
            statistics_exel = Exel.get_exel_metadata(wb)
            print("count of sheets:", wb.Sheets.Count)
            wb.Close(False)
            xl.Quit()
            return statistics_exel
        except Exception as e:
            error = traceback.format_exc()
            print('[bold red] Error:\n [/bold red]', error)
            sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)
            log.critical(f'\nFile name: {file_name}\n Error:\n{error}')
            # with io.open('./Error_exel.csv', 'w') as csvfile:
            #     error2 = [file_name, error]
            #     writer = csv.writer(csvfile, delimiter=';')
            #     writer.writerow(['File_name', 'Error'])
            #     writer.writerow(error2)
            statistics_exel = {}
            return statistics_exel

    def run_compare_exel(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for file_name in track(list_of_files, description='Comparing Exel Metadata...'):
                file_name_from = file_name.replace(f'.{extension_to}', f'.{extension_from}')
                name_to_for_test = self.preparing_file_names(file_name)
                name_from_for_test = self.preparing_file_names(file_name_from)
                if extension_to == file_name.split('.')[-1]:
                    print(f'[bold green]In test[/bold green] {file_name}')
                    print(f'[bold green]In test[/bold green] {file_name_from}')
                    self.copy(f'{custom_doc_to}{file_name}',
                              f'{path_to_temp_in_test}{name_to_for_test}')
                    self.copy(f'{custom_doc_from}{file_name_from}',
                              f'{path_to_temp_in_test}{name_from_for_test}')

                    statistics_exel_after = self.opener_exel(path_to_temp_in_test, name_from_for_test)
                    statistics_exel_before = self.opener_exel(path_to_temp_in_test, name_to_for_test)

                    if statistics_exel_after == {} or statistics_exel_before == {}:
                        print('[bold red]NOT TESTED, Statistics empty!!![/bold red]')
                        self.copy_to_not_tested(file_name,
                                                file_name_from)

                    else:
                        modified = self.dict_compare(statistics_exel_before, statistics_exel_after)
                        if modified != {}:
                            print(modified)
                            self.copy_to_errors(file_name,
                                                file_name_from)
                            modified_keys = [file_name, modified]
                            writer.writerow(modified_keys)

            self.delete(path_to_temp_in_test)
