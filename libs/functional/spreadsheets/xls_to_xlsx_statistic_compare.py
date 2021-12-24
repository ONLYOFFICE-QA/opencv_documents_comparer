# -*- coding: utf-8 -*-
import csv
import io
import os
import traceback
from multiprocessing import Process

from loguru import logger
from rich import print
from rich.progress import track
from win32com.client import Dispatch

from config import version
from libs.helpers.get_error import CheckErrors
from libs.helpers.helper import Helper

source_extension = 'xls'
converted_extension = 'xlsx'


class Excel:

    def __init__(self):
        self.helper = Helper(source_extension, converted_extension)
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.statistics_excel = None
        self.click = self.helper.click
        self.shell = Dispatch("WScript.Shell")
        logger.info(f'The {source_extension}_{converted_extension} comparison on version: {version} is running.')

    def get_excel_statistic(self, wb):
        self.statistics_excel = {
            'num_of_sheets': f'{wb.Sheets.Count}',
        }
        try:
            num_of_sheet = 1
            for sh in wb.Sheets:
                ws = wb.Worksheets(sh.Name)
                used = ws.UsedRange
                nrows = used.Row + used.Rows.Count - 1
                ncols = used.Column + used.Columns.Count - 1
                self.statistics_excel[f'{num_of_sheet}_page_name'] = sh.Name
                self.statistics_excel[f'{num_of_sheet}_nrows'] = nrows
                self.statistics_excel[f'{num_of_sheet}_ncols'] = ncols
                num_of_sheet += 1

        except Exception:
            logger.error(f'\nFailed to get full statistics excel from file: {self.helper.converted_file}\n '
                         f'statistics: {self.statistics_excel}')

    def opener_excel(self, file_name):
        error_processing = Process(target=self.check_errors.run_get_error_exel, args=(self.helper.converted_file,))
        error_processing.start()
        try:
            xl = Dispatch("Excel.Application")
            xl.Visible = False
            wb = xl.Workbooks.Open(f'{self.helper.tmp_dir_in_test}{file_name}')
            self.get_excel_statistic(wb)
            wb.Close(False)
            xl.Quit()

        except Exception:
            error = traceback.format_exc()
            logger.error(f'{error} happened while opening file: {self.helper.converted_file}')
            self.statistics_excel = None

        finally:
            os.system("taskkill /t /im  EXCEL.EXE")
            error_processing.terminate()

    def run_compare_excel_statistic(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for self.helper.converted_file in track(list_of_files,
                                                    description='[bold blue]'
                                                                'Comparing Excel Statistic...'
                                                                '[/bold blue]'):

                if self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                    self.helper.preparing_files_for_test()

                    print(f'[bold green]In test[/bold green] {self.helper.source_file} '
                          f'[bold green]and[/bold green] {self.helper.converted_file}')

                    self.opener_excel(self.helper.tmp_name_source_file)
                    source_statistics = self.statistics_excel
                    self.opener_excel(self.helper.tmp_name_converted_file)
                    converted_statistics = self.statistics_excel

                    if source_statistics is None or converted_statistics is None:
                        print("[bold red]Can't open source file, copy to untested[/bold red]")
                        self.helper.copy_to_folder(self.helper.untested_folder)

                    else:
                        modified = self.helper.dict_compare(source_statistics, converted_statistics)
                        if modified != {}:
                            print(f'[bold red]Differences: {modified}[/bold red]')
                            self.helper.copy_to_folder(self.helper.differences_statistic)

                            modified_keys = [self.helper.converted_file, modified]
                            writer.writerow(modified_keys)

            self.helper.tmp_cleaner()
