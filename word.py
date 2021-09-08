import csv
import io
import json
import os
import subprocess as sb

from rich import print
from rich.progress import track
from win32com.client import Dispatch

from Compare_Image import CompareImage
from helper import Helper
from var import *


class Word(Helper):

    def __init__(self):
        self.create_project_dirs()
        self.open_document_and_compare(os.listdir(custom_path_to_document_to),
                                       extension_from,
                                       extension_to)

    @staticmethod
    def get_word_metadata(word):
        statistics_word = {
            'num_of_sheets': f'{word.ComputeStatistics(2)}',
            'number_of_lines': f'{word.ComputeStatistics(1)}',
            'word_count': f'{word.ComputeStatistics(0)}',
            'number_of_characters_without_spaces': f'{word.ComputeStatistics(3)}',
            'number_of_characters_with_spaces': f'{word.ComputeStatistics(5)}',
            'number_of_abzad': f'{word.ComputeStatistics(4)}',
        }
        return statistics_word

    def word_opener(self, file_name, path_for_open):
        # file_name_for_test = self.preparing_files(file_name)
        # self.copy(f'{path_for_open}{file_name}',
        #                f'{path_to_temp_in_test}{file_name_for_test}')
        word = Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        # word = word.Documents.Open( f'{path_to_temp_in_test}{file_name_for_test}')
        try:
            word = word.Documents.Open(f'{path_for_open}{file_name}', None, True)
            # word.Repaginate()
            statistics_word = self.get_word_metadata(word)
        except Exception:
            print('[bold red]NOT TESTED!!![/bold red]')
            self.copy(f'{path_to_not_tested_file}{file_name}',
                      f'{path_to_not_tested_file}{file_name}')
            statistics_word = {}

        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        return statistics_word

    def open_document_and_compare(self, list_of_files, from_extension, to_extension):
        for file_name in track(list_of_files, description='[bold blue]Comparing... [/bold blue]'):
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            if to_extension == file_name.split('.')[-1]:
                print(f'[bold green]In test[/bold green] {file_name}')
                statistics_word_after = self.word_opener(file_name, custom_path_to_document_to)
                print(f'[bold green]In test[/bold green] {file_name_from}')
                statistics_word_before = self.word_opener(file_name_from, custom_path_to_document_from)

                modified = self.dict_compare(statistics_word_before, statistics_word_after)

                if modified != {}:
                    print(modified)
                    self.copy(f'{custom_path_to_document_to}{file_name}',
                              f'{path_to_errors_file}{file_name}')

                    self.copy(f'{custom_path_to_document_from}{file_name_from}',
                              f'{path_to_errors_file}{file_name_from}')

                    with io.open('report.csv', 'w', encoding="utf-8") as csvfile:
                        writer = csv.writer(csvfile, delimiter=';')
                        # writer.writerow(['file name', 'modified keys', 'values'])
                        # print('tut')
                        # print(file_name, (m for m in modified.keys()), (m[m] for m in modified.values()))
                        modified_keys = [file_name]
                        for m in modified.keys():
                            modified_keys.append(m)
                            for z in modified.values():
                                modified_keys.append(z)
                        writer.writerow(modified_keys)

                    with open(f'{path_to_errors_file}{file_name}_difference.json', 'w') as f:
                        json.dump(modified, f)
                    pass

        self.delete(path_to_temp_in_test)

    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)

