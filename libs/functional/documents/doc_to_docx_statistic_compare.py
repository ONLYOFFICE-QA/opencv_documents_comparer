# -*- coding: utf-8 -*-
import csv
import io

from loguru import logger
from rich import print
from rich.progress import track

import settings as config
from framework.word import Word


# comparison of statistical data doc-docx documents
class DocDocxStatisticsCompare(Word):
    def run_compare_statistic(self, file_array):
        self.doc_helper.terminate_process()
        logger.info(f'The {self.doc_helper.source_extension}  to {self.doc_helper.converted_extension} '
                    f'comparison of statistical data on version: {config.version} is running.')
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'num_of_sheets', 'number_of_lines', 'word_count', 'characters_without_spaces',
                             'characters_with_spaces', 'number_of_paragraph'])
            for self.doc_helper.converted_file in track(file_array, description='[blue]Comparing Word Statistic[/]\n'):
                if not self.doc_helper.converted_file.endswith((".docx", ".DOCX")):
                    continue
                self.doc_helper.preparing_files_for_compare_test()
                print(f'[bold green]In test {self.doc_helper.source_file} and {self.doc_helper.converted_file}[/]')
                if not self.get_information_about_document(self.doc_helper.tmp_source_file):
                    continue
                source_statistics = self.statistics_word
                if not self.get_information_about_document(f'{self.doc_helper.tmp_converted_file}'):
                    continue
                converted_statistics = self.statistics_word
                modified = self.doc_helper.dict_compare(source_statistics, converted_statistics)
                if modified == {}:
                    continue
                print(f'[bold red]Differences: {modified}[/]')
                self.doc_helper.copy_testing_files_to_folder(self.doc_helper.differences_statistic)
                writer.writerow(self.statistic_report_generation(modified))
                self.doc_helper.tmp_cleaner()
