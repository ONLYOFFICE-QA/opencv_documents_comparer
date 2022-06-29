# -*- coding: utf-8 -*-

import csv
import io
import json

from loguru import logger
from rich import print
from rich.progress import track

from config import version
from framework.word import Word
from libs.helpers.helper import Helper


# comparison of statistical data doc-docx documents
class DocDocxStatisticsCompare:
    def __init__(self):
        self.helper = Helper('doc', 'docx')
        self.word = Word(self.helper)
        logger.info(f'The {self.helper.source_extension}  to {self.helper.converted_extension} '
                    f'comparison of statistical data on version: {version} is running.')

    def run_compare_word_statistic(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'num_of_sheets', 'number_of_lines', 'word_count', 'characters_without_spaces',
                             'characters_with_spaces', 'number_of_paragraph'])

            for self.helper.converted_file in track(list_of_files, description='[blue]Comparing Word Statistic[/]\n'):

                if not self.helper.converted_file.endswith((".docx", ".DOCX")):
                    continue
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test {self.helper.source_file} and {self.helper.converted_file}[/]')
                if not self.word.word_opener(f'{self.helper.tmp_name_source_file}'):
                    continue
                source_statistics = self.word.statistics_word
                if not self.word.word_opener(f'{self.helper.tmp_name_converted_file}'):
                    continue
                converted_statistics = self.word.statistics_word
                modified = self.helper.dict_compare(source_statistics, converted_statistics)

                if modified == {}:
                    continue
                print(f'[bold red]Differences: {modified}[/]')
                self.helper.copy_to_folder(self.helper.differences_statistic)
                writer.writerow(self.word.statistic_report_generation(modified))
            self.helper.tmp_cleaner()
