# -*- coding: utf-8 -*-
from loguru import logger
from rich import print
from rich.progress import track
import csv
import io
import configuration as config

from framework.excel import Excel


class StatisticCompare(Excel):
    def run_compare_statistic(self, files_array):
        logger.info(f'The {self.doc_helper.source_extension} to {self.doc_helper.converted_extension} '
                    f'comparison of statistical data on version: {config.version} is running.')
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for self.doc_helper.converted_file in track(files_array, description='[blue]Comparing Excel Statistic[/]'):
                if not self.doc_helper.converted_file.endswith((".xlsx", ".XLSX")):
                    continue
                self.doc_helper.preparing_files_for_compare_test()
                print(f'[bold green]In test: {self.doc_helper.source_file} and {self.doc_helper.converted_file}[/]')
                if not self.get_information_about_table(self.doc_helper.tmp_source_file):
                    continue
                source_statistics = self.statistics_excel
                if not self.get_information_about_table(self.doc_helper.tmp_converted_file):
                    continue
                converted_statistics = self.statistics_excel
                modified = self.doc_helper.dict_compare(source_statistics, converted_statistics)
                if modified == {}:
                    continue
                print(f'[bold red]Differences: {modified}[/]')
                self.doc_helper.copy_testing_files_to_folder(self.doc_helper.differences_statistic)
                modified_keys = [self.doc_helper.converted_file, modified]
                writer.writerow(modified_keys)
                self.doc_helper.tmp_cleaner()
