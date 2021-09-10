import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print

from libs.Compare_Image import CompareImage
from libs.helper import Helper
from libs.word import Word
from var import *


class WordCompareImg(Helper):

    def __init__(self):
        self.create_project_dirs()
        self.copy_for_test()
        self.coordinate = []
        self.run_compare_word(list_file_names_doc_from_compare)

    def get_coord_word(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'OpusApp':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 494, 30, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def get_screenshots(self, file_name_for_screen, path_to_save_screen, num_of_sheets):
        print(f'[bold green]In test[/bold green] {file_name_for_screen}')
        Word.run(path_to_folder_for_test, file_name_for_screen, 'WINWORD.EXE')
        sleep(wait_for_open)
        win32gui.EnumWindows(self.get_coord_word, self.coordinate)
        page_num = 1
        for page in range(int(num_of_sheets)):
            CompareImage.grab_coordinate(path_to_save_screen, file_name_for_screen, page_num, self.coordinate[0])
            # pg.click()
            pg.press('pgdn')
            sleep(wait_for_press)
            page_num += 1
        # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        sb.call(["taskkill", "/IM", "WINWORD.EXE"])

    def run_compare_word(self, list_of_files, from_extension=extension_from, to_extension=extension_to):
        for file_name in list_of_files:
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            if to_extension == file_name.split('.')[-1]:
                num_of_sheets = Word.word_opener(path_to_folder_for_test, file_name_from)
                print(num_of_sheets['num_of_sheets'])
                if num_of_sheets != {}:
                    self.get_screenshots(file_name, path_to_tmpimg_after_conversion,
                                         num_of_sheets['num_of_sheets'])
                    self.get_screenshots(file_name_from, path_to_tmpimg_befor_conversion,
                                         num_of_sheets['num_of_sheets'])
                    CompareImage()

        self.delete(path_to_temp_in_test)
