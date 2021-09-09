from time import sleep

from libs.Compare_Image import CompareImage
from libs.helper import Helper
from libs.word import Word
from var import *
import pyautogui as pg
import subprocess as sb

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
                num_of_sheets = Word.word_opener(path_to_folder_for_test, file_name_from)
                print(num_of_sheets['num_of_sheets'])
                if num_of_sheets != {}:
                    self.get_screenshots(file_name, path_to_tmpimg_after_conversion, num_of_sheets['num_of_sheets'])
                    self.get_screenshots(file_name_from, path_to_tmpimg_befor_conversion,
                                         num_of_sheets['num_of_sheets'])
                    CompareImage()

        self.delete(path_to_temp_in_test)

