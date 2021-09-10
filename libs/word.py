import csv
import io
import json

from rich import print
from rich.progress import track
from win32com.client import Dispatch

from libs.helper import Helper
from var import *


class Word(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        self.open_document_and_compare(list_of_files,
                                       extension_from,
                                       extension_to)
        # self.open_document_and_compare(list_file_names_doc_from_compare,
        #                                extension_from,
        #                                extension_to)

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

    def word_opener(path_for_open, file_name):
        word_app = Dispatch('Word.Application')
        word_app.Visible = False
        # word_app.DisplayAlerts = False
        try:
            word_app = word_app.Documents.Open(f'{path_for_open}{file_name}', None, True)
            statistics_word = Word.get_word_metadata(word_app)
        except Exception:
            print('[bold red]NOT TESTED!!![/bold red]')
            statistics_word = {}
        word_app.Close(False)
        # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        # sb.call(["taskkill", "/IM", "WINWORD.EXE", "/T"])
        return statistics_word

    def open_document_and_compare(self, list_of_files, from_extension=extension_from, to_extension=extension_to):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            for file_name in track(list_of_files, description='[bold blue]Comparing... [/bold blue]'):
                file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
                name_for_test = self.preparing_file_names(file_name)
                name_from_for_test = self.preparing_file_names(file_name_from)
                if to_extension == file_name.split('.')[-1]:
                    print(f'[bold green]In test[/bold green] {file_name}')
                    self.copy(f'{custom_doc_to}{file_name}',
                              f'{path_to_temp_in_test}{name_for_test}')
                    statistics_word_after = Word.word_opener(path_to_temp_in_test,
                                                             name_for_test)

                    print(f'[bold green]In test[/bold green] {file_name_from}')
                    self.copy(f'{custom_doc_from}{file_name_from}',
                              f'{path_to_temp_in_test}{name_from_for_test}')
                    statistics_word_before = Word.word_opener(path_to_temp_in_test,
                                                              name_from_for_test)

                    if statistics_word_after == {} or statistics_word_before == {}:
                        print('[bold red]NOT TESTED, Statistics empty!!![/bold red]')
                        self.copy_to_not_tested(file_name,
                                                file_name_from)

                    else:
                        modified = self.dict_compare(statistics_word_before, statistics_word_after)

                        if modified != {}:
                            print(modified)
                            self.copy_to_errors(file_name,
                                                file_name_from)

                            writer = csv.writer(csvfile, delimiter=';')
                            modified_keys = [file_name]
                            for m in modified:
                                modified_keys.append(m)
                                modified_keys.append(modified[m])
                            writer.writerow(modified_keys)

                            with open(f'{path_to_errors_file}{file_name}_difference.json', 'w') as f:
                                json.dump(modified, f)


        self.delete(path_to_temp_in_test)


