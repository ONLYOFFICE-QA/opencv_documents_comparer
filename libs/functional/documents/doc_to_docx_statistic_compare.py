# -*- coding: utf-8 -*-
import csv
import io
import json
import os
from multiprocessing import Process

from loguru import logger
from rich import print
from rich.progress import track
from win32com.client import Dispatch

from config import version
from libs.helpers.get_error import CheckErrors
from libs.helpers.helper import Helper

source_extension = 'doc'
converted_extension = 'docx'


class Word:

    def __init__(self):
        self.helper = Helper(source_extension, converted_extension)
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.statistics_word = None
        self.shell = Dispatch("WScript.Shell")
        self.click = self.helper.click
        logger.info(f'The {source_extension}_{converted_extension} comparison on version: {version} is running.')

    def get_word_statistic(self, word_app):
        try:
            self.statistics_word = {
                'num_of_sheets': f'{word_app.ComputeStatistics(2)}',
                'number_of_lines': f'{word_app.ComputeStatistics(1)}',
                'word_count': f'{word_app.ComputeStatistics(0)}',
                'number_of_characters_without_spaces': f'{word_app.ComputeStatistics(3)}',
                'number_of_characters_with_spaces': f'{word_app.ComputeStatistics(5)}',
                'number_of_paragraph': f'{word_app.ComputeStatistics(4)}',
            }

        except Exception:
            logger.error(f'Exception while getting statistics, {self.helper.converted_file}')
            self.statistics_word = None

    def word_opener(self, file_name_for_open):
        error_processing = Process(target=self.check_errors.run_get_errors_word, args=(self.helper.converted_file,))
        error_processing.start()
        word_app = Dispatch('Word.Application')
        word_app.Visible = False
        try:
            word_app = word_app.Documents.Open(f'{self.helper.tmp_dir_in_test}{file_name_for_open}', None, True)
            self.get_word_statistic(word_app)
            word_app.Close(False)

        except Exception:
            logger.exception(f'Exception while opening: {self.helper.converted_file}')

        finally:
            os.system("taskkill /t /im  WINWORD.EXE")
            error_processing.terminate()

    def run_compare_word_statistic(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'num_of_sheets', 'number_of_lines', 'word_count', 'characters_without_spaces',
                             'characters_with_spaces', 'number_of_paragraph'])

            for self.helper.converted_file in track(list_of_files,
                                                    description='[bold blue]Comparing Word Statistic...[/bold blue]\n'):

                if self.helper.converted_file.endswith((".docx", ".DOCX")):
                    self.helper.preparing_files_for_test()

                    print(f'[bold green]In test[/bold green] {self.helper.source_file} '
                          f'[bold green]and[/bold green] {self.helper.converted_file}')

                    self.word_opener(f'{self.helper.tmp_name_source_file}')
                    source_statistics = self.statistics_word
                    self.word_opener(f'{self.helper.tmp_name_converted_file}')
                    converted_statistics = self.statistics_word

                    if source_statistics is None or converted_statistics is None:
                        logger.error(f"Can't open {self.helper.source_file} "
                                     f"or {self.helper.converted_file}. Copied files"
                                     "to 'failed_to_open_file'")

                        self.helper.copy_to_folder(self.helper.failed_source)

                    else:
                        modified = self.helper.dict_compare(source_statistics, converted_statistics)

                        if modified != {}:
                            print(f'[bold red]Differences: {modified}[/bold red]')
                            self.helper.copy_to_folder(self.helper.differences_statistic)

                            # report generation
                            modified_keys = [self.helper.converted_file]
                            for key in modified:
                                modified_keys.append(modified['num_of_sheets']) if key == 'num_of_sheets' \
                                    else modified_keys.append(' ')
                                modified_keys.append(modified['number_of_lines']) if key == 'number_of_lines' \
                                    else modified_keys.append(' ')
                                modified_keys.append(modified['word_count']) if key == 'word_count' \
                                    else modified_keys.append(' ')
                                modified_keys.append(modified[
                                                         'number_of_characters_without_spaces']) \
                                    if key == 'number_of_characters_without_spaces' \
                                    else modified_keys.append(' ')
                                modified_keys.append(modified[
                                                         'number_of_characters_with_spaces']) \
                                    if key == 'number_of_characters_with_spaces' \
                                    else modified_keys.append(' ')
                                modified_keys.append(modified['number_of_paragraph']) if key == 'number_of_paragraph' \
                                    else modified_keys.append(' ')

                            writer.writerow(modified_keys)

                            # Saving differences in json
                            with open(
                                    f'{self.helper.differences_statistic}{self.helper.converted_file}_difference.json',
                                    'w') as f:
                                json.dump(modified, f)

            self.helper.tmp_cleaner()
