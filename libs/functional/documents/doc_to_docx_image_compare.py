# -*- coding: utf-8 -*-
import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print

from config import *
from libs.functional.documents.doc_to_docx_statistic_compare import Word
from libs.helpers.compare_image import CompareImage
from libs.helpers.get_error import check_word

source_extension = 'doc'
converted_extension = 'docx'


class WordCompareImg(Word):

    # gets the coordinates of the window
    # sets the size and position of the window
    def get_coord_word(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'OpusApp':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 0, 0, 2000, 1400, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def check_error_word(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' \
                    or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword' \
                    or win32gui.GetClassName(hwnd) == 'ThunderDFrame' \
                    or win32gui.GetClassName(hwnd) == 'NUIDialog':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)

                self.errors.clear()
                self.errors.append(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetWindowText(hwnd))

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, tmp_file_name, path_to_save_screen, num_of_sheets):
        self.helper.run(self.helper.tmp_dir_in_test, tmp_file_name, 'WINWORD.EXE')
        sleep(wait_for_opening)
        win32gui.EnumWindows(self.get_coord_word, self.coordinate)
        for num_errors in range(5):
            win32gui.EnumWindows(self.check_error_word, self.errors)
            sleep(0.2)
            if self.errors:
                check_word()

        coordinate = self.coordinate[0]
        coordinate = (coordinate[0],
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 20)
        page_num = 1
        for page in range(int(num_of_sheets)):
            CompareImage.grab_coordinate(path_to_save_screen, tmp_file_name, page_num, coordinate)
            # pg.click()
            pg.press('pgdn')
            sleep(wait_for_press)
            page_num += 1
        sb.call(["taskkill", "/IM", "WINWORD.EXE"], shell=True)
        sb.call(["taskkill", "/IM", "WINWORD.EXE"], shell=True)

    def run_compare_word(self, list_of_files):
        for converted_file in list_of_files:
            if converted_file.endswith((".docx", ".DOCX")):
                if converted_file == 'Integrated ICT for development' \
                                     ' program Recommendations for ' \
                                     'USAID Macedonia focused on ' \
                                     'education and workforce training.docx':
                    converted_file = 'IntegratedICTfordevelopment_renamed.docx'

                source_file, tmp_name_converted_file, \
                tmp_name_source_file, tmp_name = self.helper.preparing_files_for_test(converted_file,
                                                                                      converted_extension,
                                                                                      source_extension)
                print(f'[bold green]In test[/bold green] {converted_file}')
                num_of_sheets = Word.word_opener(f'{self.helper.tmp_dir_in_test}{tmp_name}')

                if num_of_sheets != {}:
                    num_of_sheets = num_of_sheets['num_of_sheets']
                    print(f"Number of pages: {num_of_sheets}")
                    self.get_screenshots(tmp_name_converted_file,
                                         self.helper.tmp_dir_converted_image,
                                         num_of_sheets)

                    self.get_screenshots(tmp_name_source_file,
                                         self.helper.tmp_dir_source_image,
                                         num_of_sheets)

                    CompareImage(converted_file, self.helper)

        self.helper.delete(f'{self.helper.tmp_dir_in_test}*')
