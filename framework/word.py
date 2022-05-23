# -*- coding: utf-8 -*-

import os
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from win32com.client import Dispatch

from config import wait_for_opening, wait_for_press
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


# methods for working with Word
class Word:

    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.statistics_word = None
        self.shell = Dispatch("WScript.Shell")
        self.click = self.helper.click

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

    def check_error_for_opener(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' \
                    and win32gui.GetWindowText(hwnd) == "Microsoft Word":
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.check_errors.errors.clear()
                self.check_errors.errors.append(win32gui.GetClassName(hwnd))
                self.check_errors.errors.append(win32gui.GetWindowText(hwnd))

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

    def open_word_with_cmd_for_opener(self, file_name):
        self.helper.run(self.helper.tmp_dir_in_test, file_name, 'WINWORD.EXE')
        sleep(wait_for_opening)
        win32gui.EnumWindows(self.check_error_for_opener, self.check_errors.errors)

    def open_word_with_cmd(self, file_name):
        self.helper.run(self.helper.tmp_dir_in_test, file_name, 'WINWORD.EXE')
        sleep(wait_for_opening)
        self.check_error()

    def close_word_with_cmd(self):
        pg.hotkey('ctrl', 'z')
        pg.press('esc')
        os.system("taskkill /t /im  WINWORD.EXE")
        sleep(0.2)
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == 'NUIDialog' \
                and self.check_errors.errors[1] == "Microsoft Word":
            pg.press('right')
            pg.press('enter')
            self.check_errors.errors.clear()

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, path_to_save_screen):

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
