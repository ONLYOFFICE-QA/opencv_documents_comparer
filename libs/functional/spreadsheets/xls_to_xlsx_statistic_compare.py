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
            for self.helper.converted_file in track(list_of_files, description='[blue]Comparing Excel Statistic[/]'):
                if not self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                    continue
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test: {self.helper.source_file} and {self.helper.converted_file}[/]')
                if not self.excel.opener_excel(self.helper.tmp_name_source_file):
                    continue
                source_statistics = self.excel.statistics_excel
                if not self.excel.opener_excel(self.helper.tmp_name_converted_file):
                    continue
                converted_statistics = self.excel.statistics_excel

                modified = self.helper.dict_compare(source_statistics, converted_statistics)
                if modified == {}:
                    continue
                print(f'[bold red]Differences: {modified}[/]')
                self.helper.copy_to_folder(self.helper.differences_statistic)
                modified_keys = [self.helper.converted_file, modified]
                writer.writerow(modified_keys)
            self.helper.tmp_cleaner()
