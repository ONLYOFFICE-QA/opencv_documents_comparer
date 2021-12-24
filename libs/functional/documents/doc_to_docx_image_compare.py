# -*- coding: utf-8 -*-
import os
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from rich import print

from config import *
from libs.functional.documents.doc_to_docx_statistic_compare import Word
from libs.helpers.compare_image import CompareImage

source_extension = 'doc'
converted_extension = 'docx'


class WordCompareImg(Word):

    # gets the coordinates of the window
    # sets the size and position of the window
    def get_coord_word(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'OpusApp':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 0, 0, 2000, 1400, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def check_error(self):
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
        if self.check_errors.errors:
            self.check_errors.errors.clear()
            error_processing = Process(target=self.check_errors.run_get_errors_word, args=(self.helper.converted_file,))
            error_processing.start()
            sleep(7)
            error_processing.terminate()

    def prepare_word_windows(self):
        self.click('libs/image_templates/word/layout.png')
        sleep(0.3)
        self.click('libs/image_templates/word/transfers.png')
        pg.press('down', interval=0.1)
        pg.press('enter')
        self.click('libs/image_templates/powerpoint/view.png')
        sleep(0.3)
        self.click('libs/image_templates/word/one_page.png')
        self.click('libs/image_templates/word/resolution100.png')
        pg.moveTo(100, 0)
        sleep(1)

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, path_to_save_screen, file_name):
        self.helper.run(self.helper.tmp_dir_in_test, file_name, 'WINWORD.EXE')
        sleep(wait_for_opening)
        self.check_error()

        win32gui.EnumWindows(self.get_coord_word, self.coordinate)
        coordinate = self.coordinate[0]
        coordinate = (coordinate[0],
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 20)

        self.prepare_word_windows()
        page_num = 1
        for page in range(int(self.statistics_word['num_of_sheets'])):
            CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
            pg.press('pgdn')
            sleep(wait_for_press)
            page_num += 1
        pg.hotkey('ctrl', 'z')
        os.system("taskkill /t /im  WINWORD.EXE")

    def run_compare_word(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".docx", ".DOCX")):
                self.helper.preparing_files_for_test()

                if self.helper.converted_file == 'Integrated ICT for development' \
                                                 ' program Recommendations for ' \
                                                 'USAID Macedonia focused on ' \
                                                 'education and workforce training.docx':
                    self.helper.converted_file = 'IntegratedICTfordevelopment_renamed.docx'

                print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
                self.word_opener(self.helper.tmp_name)

                if self.statistics_word is not None:
                    print(f"[bold blue]Number of pages:[/bold blue] {self.statistics_word['num_of_sheets']}")
                    self.get_screenshots(self.helper.tmp_dir_converted_image, self.helper.tmp_name_converted_file)

                    print(f'[bold green]In test[/bold green] {self.helper.source_file}')
                    self.get_screenshots(self.helper.tmp_dir_source_image, self.helper.tmp_name_source_file)

                    CompareImage(self.helper)

                elif self.statistics_word is None:
                    logger.error(f"Can't get number of pages in {self.helper.source_file}. Copied files "
                                 "to 'failed_to_open_file'")
                    self.helper.copy_to_folder(self.helper.failed_source)
                else:
                    logger.debug(f"Debug. Statistics word: {self.statistics_word}")

            self.helper.tmp_cleaner()
