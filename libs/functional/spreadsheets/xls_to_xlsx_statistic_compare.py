# -*- coding: utf-8 -*-
import csv
import io

from loguru import logger
from rich import print
from rich.progress import track

from config import version
from framework.excel import Excel
from libs.helpers.helper import Helper


class StatisticCompare:

    def __init__(self):
        self.helper = Helper('xls', 'xlsx')
        self.excel = Excel(self.helper)
        logger.info(f'The {self.helper.source_extension} to {self.helper.converted_extension} '
                    f'comparison of statistical data on version: {version} is running.')

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

                    self.excel.opener_excel(self.helper.tmp_name_source_file)
                    source_statistics = self.excel.statistics_excel
                    self.excel.opener_excel(self.helper.tmp_name_converted_file)
                    converted_statistics = self.excel.statistics_excel

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
