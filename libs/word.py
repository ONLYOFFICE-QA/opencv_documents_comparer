import csv
import io
import json
import subprocess as sb
from time import sleep

from rich import print
from rich.progress import track
from win32com.client import Dispatch
import pyautogui as pg

from libs.Compare_Image import CompareImage
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
        word = Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        try:
            word = word.Documents.Open(f'{path_for_open}{file_name}', None, True)
            # word.Repaginate()
            statistics_word = Word.get_word_metadata(word)
        except Exception:
            print('[bold red]NOT TESTED!!![/bold red]')
            Word.copy(f'{path_for_open}{file_name}',
                      f'{path_to_not_tested_file}{file_name}')
            statistics_word = {}

        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        return statistics_word

    def open_document_and_compare(self, list_of_files, from_extension=extension_from, to_extension=extension_to):
        with io.open('../report.csv', 'w', encoding="utf-8") as csvfile:
            for file_name in track(list_of_files, description='[bold blue]Comparing... [/bold blue]'):
                file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
                if to_extension == file_name.split('.')[-1]:
                    print(f'[bold green]In test[/bold green] {file_name}')
                    statistics_word_after = Word.word_opener(custom_path_to_document_to,
                                                             file_name)
                    print(f'[bold green]In test[/bold green] {file_name_from}')
                    statistics_word_before = Word.word_opener(custom_path_to_document_from,
                                                              file_name_from)

                    modified = self.dict_compare(statistics_word_before, statistics_word_after)

                    if modified != {}:
                        print(modified)
                        self.copy(f'{custom_path_to_document_to}{file_name}',
                                  f'{path_to_errors_file}{file_name}')

                        self.copy(f'{custom_path_to_document_from}{file_name_from}',
                                  f'{path_to_errors_file}{file_name_from}')

                        writer = csv.writer(csvfile, delimiter=';')
                        modified_keys = [file_name]
                        for m in modified:
                            modified_keys.append(m)
                            modified_keys.append(modified[m])
                        writer.writerow(modified_keys)

                        with open(f'{path_to_errors_file}{file_name}_difference.json', 'w') as f:
                            json.dump(modified, f)
                        pass

        self.delete(path_to_temp_in_test)


class WordCompareImg(Helper):
    def __init__(self):
        self.create_project_dirs()
        self.copy_for_test()
        self.run_compare(list_file_names_doc_from_compare)

    @staticmethod
    def get_screenshots(file_name, path_to_save_screen, num_of_sheets):
        print(f'[bold green]In test[/bold green] {file_name}')
        Word.run(path_to_folder_for_test, file_name, 'WINWORD.EXE')
        sleep(wait_for_open)
        page_num = 1
        for page in range(int(num_of_sheets)):
            CompareImage.grab(path_to_save_screen, file_name, page_num)
            pg.click()
            pg.press('pgdn')
            sleep(wait_for_press)
            page_num += 1
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)

    def run_compare(self, list_of_files, from_extension=extension_from, to_extension=extension_to):
        for file_name in list_of_files:
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            if to_extension == file_name.split('.')[-1]:
                num_of_sheets = Word.word_opener(path_to_folder_for_test, file_name)
                self.get_screenshots(file_name, path_to_tmpimg_after_conversion, num_of_sheets['num_of_sheets'])
                self.get_screenshots(file_name_from, path_to_tmpimg_befor_conversion, num_of_sheets['num_of_sheets'])
                CompareImage()

        self.delete(path_to_temp_in_test)

    pass
