# -*- coding: utf-8 -*-

import csv
import io
import json

from loguru import logger
from rich import print
from rich.progress import track

from libs.functional.documents.word import Word
from libs.helpers.helper import Helper


# comparison of statistical data doc-docx documents
class DocDocxStatisticsCompare:

    def __init__(self):
        self.source_extension = 'doc'
        self.converted_extension = 'docx'
        self.helper = Helper(self.source_extension, self.converted_extension)
        self.word = Word(self.helper)
        pass

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

                    self.word.word_opener(f'{self.helper.tmp_name_source_file}')
                    source_statistics = self.word.statistics_word
                    self.word.word_opener(f'{self.helper.tmp_name_converted_file}')
                    converted_statistics = self.word.statistics_word

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
