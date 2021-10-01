import csv
import io
import subprocess as sb
import traceback

import win32con
import win32gui
from rich import print
from rich.progress import track
from win32com.client import Dispatch

from libs.helpers.helper import Helper
from libs.helpers.logger import *
from variables import *

extension_source = 'xls'
extension_converted = 'xlsx'


class Excel(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        self.delete(f'{tmp_dir_in_test}*')
        self.delete(f'{tmp_dir_converted_image}*')
        self.delete(f'{tmp_dir_source_image}*')
        self.coordinate = []
        self.errors = []
        self.run_compare_exel(list_of_files)

    @staticmethod
    def get_exel_statistic(wb):
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

    @staticmethod
    def opener_exel(path_for_open, file_name):
        try:
            xl = Dispatch("Excel.Application")
            xl.Visible = False  # otherwise excel is hidden
            wb = xl.Workbooks.Open(f'{path_for_open}{file_name}', ReadOnly=True)
            statistics_exel = Excel.get_exel_statistic(wb)
            wb.Close(False)
            xl.Quit()
            return statistics_exel

        except Exception as e:
            error = traceback.format_exc()
            print('[bold red] Error:\n [/bold red]', error)
            sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)
            log.critical(f'\nFile name: {file_name}\n Error:\n{error}')
            statistics_exel = {}
            return statistics_exel

    def run_compare_exel(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for converted_file in track(list_of_files, description='Comparing Excel Metadata...'):
                if converted_file.endswith((".xlsx", ".XLSX")):
                    source_file, tmp_name_converted_file, \
                    tmp_name_source_file, tmp_name = self.preparing_files_for_test(converted_file,
                                                                                   converted_extension,
                                                                                   source_extension)

                    print(f'[bold green]In test[/bold green] {source_file}')
                    statistics_exel_after = self.opener_exel(tmp_dir_in_test, tmp_name_source_file)
                    print(f'[bold green]In test[/bold green] {converted_file}')
                    statistics_exel_before = self.opener_exel(tmp_dir_in_test, tmp_name_converted_file)

                    if statistics_exel_after == {} or statistics_exel_before == {}:
                        print("[bold red]Can't open source file, copy to untested[/bold red]")
                        self.copy_to_folder(converted_file,
                                            source_file,
                                            untested_folder)

                    else:
                        modified = self.dict_compare(statistics_exel_before, statistics_exel_after)
                        if modified != {}:
                            print(modified)
                            self.copy_to_folder(converted_file,
                                                source_file,
                                                differences_statistic)
                            modified_keys = [converted_file, modified]
                            writer.writerow(modified_keys)

            self.delete(tmp_dir_in_test)
