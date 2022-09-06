# -*- coding: utf-8 -*-
import os
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from win32com.client import Dispatch
from rich import print

from config import wait_for_opening
from framework.telegram import Telegram
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


# methods for working with Word
class Word:
    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.click = self.helper.click
        self.statistics_word = None
        self.windows_handler_number = None

    def check_error_for_opener(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd) \
                and win32gui.GetClassName(hwnd) == "#32770" \
                and win32gui.GetWindowText(hwnd) == "Microsoft Word":
            self.check_errors.errors.clear()
            win32gui.SetForegroundWindow(hwnd)
            self.check_errors.errors.append("#32770")
            self.check_errors.errors.append("Microsoft Word")

    # sets the size and position of the window
    def set_windows_size_word(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2000, 1400, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    def prepare_document_for_test(self):
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
        sleep(0.5)

    def open_word_with_cmd(self, file_name):
        self.check_errors.errors.clear()
        self.helper.run(self.helper.tmp_dir_in_test, file_name, 'WINWORD.EXE')
        self.waiting_for_opening_word()

    def check_open_word(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) != "":
                self.windows_handler_number = hwnd
            elif win32gui.GetClassName(hwnd) == "#32770" and win32gui.GetWindowText(hwnd) == "Microsoft Word":
                logger.debug(f"document recovery {self.helper.converted_file}")
                win32gui.SetForegroundWindow(hwnd)
                pg.press('right')
                pg.press('enter')

    def waiting_for_opening_word(self):
        self.windows_handler_number = None
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_word, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open file {self.helper.converted_file}")
                self.helper.copy_to_folder(self.helper.too_long_to_open_files)
                break

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error_for_opener, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == "#32770"\
                and self.check_errors.errors[1] == "Microsoft Word":
            logger.error(f"'an error has occurred while opening' when opening file: {self.helper.converted_file}.")
            pg.press('esc', presses=3, interval=0.2)
            self.helper.copy_to_folder(self.helper.opener_errors)
            self.check_errors.errors.clear()
            return False
        elif not self.check_errors.errors:
            return True
        else:
            logger.debug(f"New Error\nError message: {self.check_errors.errors}\nFile: {self.helper.converted_file}")
            self.helper.copy_to_folder(self.helper.failed_source)
            self.check_errors.errors.clear()
            return False

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == "NUIDialog" \
                and self.check_errors.errors[1] == "Microsoft Word":
            print(f'Save file: {self.helper.converted_file}')
            pg.press('right')
            pg.press('enter')
        elif self.check_errors.errors \
                and self.check_errors.errors[0] == "#32770" \
                and self.check_errors.errors[1] == "Microsoft Word":
            logger.debug(f"operation aborted {self.helper.converted_file}")
            pg.press('enter')
        self.check_errors.errors.clear()

    def events_handler_when_opening(self):
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
        if self.check_errors.errors:
            self.check_errors.errors.clear()
            error_processing = Process(target=self.check_errors.run_get_errors_word, args=(self.helper.converted_file,))
            error_processing.start()
            sleep(7)
            error_processing.terminate()

    def close_word_with_cmd(self):
        pg.hotkey('ctrl', 'z')
        pg.press('esc')
        os.system("taskkill /t /im  WINWORD.EXE")
        sleep(0.2)
        self.events_handler_when_closing()

    # gets the coordinates of the window
    def get_coordinate_word(self):
        coordinate = [win32gui.GetWindowRect(self.windows_handler_number)]
        coordinate = coordinate[0]
        coordinate = (coordinate[0],
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 20)
        return coordinate

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, path_to_save_screen):
        if win32gui.IsWindow(self.windows_handler_number):
            self.set_windows_size_word()
            coordinate = self.get_coordinate_word()
            self.prepare_document_for_test()
            page_num = 1
            for page in range(int(self.statistics_word['num_of_sheets'])):
                CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {self.helper.converted_file}'
            Telegram.send_message(massage)
            logger.error(massage)

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
        except Exception as e:
            Telegram.send_message(f'Exception while getting statistics, {self.helper.converted_file}\nException: {e}')
            logger.exception(f'Exception while getting statistics, {self.helper.converted_file}\nException: {e}')
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
            print(f"[bold blue]Number of pages:[/] {self.statistics_word['num_of_sheets']}")
            return True
        except Exception as e:
            logger.exception(f"Can't get number of pages in {self.helper.source_file}. Exception: {e}")
            self.helper.copy_to_folder(self.helper.failed_source)
            return False
        finally:
            os.system("taskkill /t /im  WINWORD.EXE")
            error_processing.terminate()

    def statistic_report_generation(self, modified):
        modified_keys = [self.helper.converted_file]
        for key in modified:
            modified_keys.append(modified['num_of_sheets']) if key == 'num_of_sheets' else modified_keys.append(' ')
            modified_keys.append(modified['number_of_lines']) if key == 'number_of_lines' else modified_keys.append(' ')
            modified_keys.append(modified['word_count']) if key == 'word_count' else modified_keys.append(' ')
            modified_keys.append(modified['number_of_characters_without_spaces']) \
                if key == 'number_of_characters_without_spaces' else modified_keys.append(' ')
            modified_keys.append(modified['number_of_characters_with_spaces']) \
                if key == 'number_of_characters_with_spaces' else modified_keys.append(' ')
            modified_keys.append(modified['number_of_paragraph']) \
                if key == 'number_of_paragraph' else modified_keys.append(' ')
        return modified_keys
