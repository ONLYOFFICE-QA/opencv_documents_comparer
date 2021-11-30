# -*- coding: utf-8 -*-
import csv
import io
import subprocess as sb
import traceback
from multiprocessing import Process

from loguru import logger
from rich import print
from rich.progress import track
from win32com.client import Dispatch

from libs.helpers.get_error import CheckErrors
from libs.helpers.helper import Helper

source_extension = 'xls'
converted_extension = 'xlsx'


class Excel:

    def __init__(self):
        self.check_errors = CheckErrors()
        self.helper = Helper(source_extension, converted_extension)
        self.coordinate = []
        self.click = self.helper.click
        logger.add('./logs/xls_xlsx.log',
                   format="{time} {level} {message}",
                   level="DEBUG",
                   mode='w')

    @staticmethod
    def get_excel_statistic(wb, file_name_for_log):
        statistics_exel = {
            'num_of_sheets': f'{wb.Sheets.Count}',
        }
        try:
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

        except Exception:
            logger.exception(f'Failed to get full statistics excel from file: {file_name_for_log}')
            return statistics_exel

    def opener_excel(self, path_for_open, file_name, file_name_for_log):
        error_processing = Process(target=self.check_errors.run_get_error_exel, args=(file_name_for_log,))
        error_processing.start()
        try:
            xl = Dispatch("Excel.Application")
            xl.Visible = False  # otherwise excel is hidden
            wb = xl.Workbooks.Open(f'{path_for_open}{file_name}')
            statistics_excel = Excel.get_excel_statistic(wb, file_name_for_log)
            wb.Close(False)
            xl.Quit()
            return statistics_excel

        except Exception as e:
            error = traceback.format_exc()
            logger.exception(f'{error} happened while opening file: {file_name_for_log}')
            sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)
            statistics_excel = {}
            return statistics_excel

        finally:
            error_processing.terminate()

    def run_compare_excel_statistic(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for converted_file in track(list_of_files,
                                        description='[bold blue]Comparing Excel Statistic...[/bold blue]\n'):
                if converted_file.endswith((".xlsx", ".XLSX")):
                    source_file, tmp_name_converted_file, \
                    tmp_name_source_file, tmp_name = self.helper.preparing_files_for_test(converted_file,
                                                                                          converted_extension,
                                                                                          source_extension)

                    print(f'[bold green]In test[/bold green] {source_file} '
                          f'[bold green]and[/bold green] {converted_file}')
                    source_statistics = self.opener_excel(self.helper.tmp_dir_in_test, tmp_name_source_file,
                                                          source_file)
                    converted_statistics = self.opener_excel(self.helper.tmp_dir_in_test, tmp_name_converted_file,
                                                             converted_file)

                    if source_statistics == {} or converted_statistics == {}:
                        print("[bold red]Can't open source file, copy to untested[/bold red]")
                        self.helper.copy_to_folder(converted_file,
                                                   source_file,
                                                   self.helper.untested_folder)

                    else:
                        modified = self.helper.dict_compare(source_statistics, converted_statistics)
                        if modified != {}:
                            print(f'[bold red]Differences: {modified}[/bold red]')
                            self.helper.copy_to_folder(converted_file,
                                                       source_file,
                                                       self.helper.differences_statistic)

                            modified_keys = [converted_file, modified]
                            writer.writerow(modified_keys)

            self.helper.tmp_cleaner()
