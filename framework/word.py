# -*- coding: utf-8 -*-
import subprocess as sb
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from rich import print
from win32com.client import Dispatch

import settings as config
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.key_actions import KeyActions
from framework.compare_image import CompareImage
from framework.telegram import Telegram


# methods for working with Word
class Word:
    def __init__(self):
        self.doc_helper = ProjectConfig.DOC_ACTIONS
        self.statistics_word = None
        self.windows_handler_number = None
        self.files_with_errors_when_opening = []
        self.num_of_page = ''

    @staticmethod
    def get_errors():
        errors = []

        def get_windows_title(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
                if class_name in ['#32770', 'bosa_sdm_msword', 'ThunderDFrame', 'NUIDialog']:
                    win32gui.SetForegroundWindow(hwnd)
                    errors.append(class_name)
                    errors.append(window_text)

        win32gui.EnumWindows(get_windows_title, errors)
        return errors

    def errors_handler_for_thread(self):
        while True:
            errors = self.get_errors()
            if errors:
                match errors:
                    case ['#32770', 'Microsoft Word'] | \
                         ['NUIDialog', 'Извещение системы безопасности Microsoft Word']:
                        pg.press('left')
                        pg.press('enter')
                    case ['#32770', 'Microsoft Visual Basic for Applications'] | \
                         ['bosa_sdm_msword', 'Преобразование файла'] | \
                         ['NUIDialog', 'Microsoft Word']:
                        pg.press('enter')
                    case ['bosa_sdm_msword', 'Пароль']:
                        pg.press('tab')
                        pg.press('enter')
                    case ['#32770', 'Удаление нескольких элементов']:
                        print(errors)
                    case ['#32770', 'Сохранить как']:
                        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
                    case ['bosa_sdm_msword', 'Показать исправления']:
                        sleep(2)
                        pg.press('tab', presses=3)
                        pg.press('enter')
                    case _:
                        message = f"New Event {errors} happened while opening: {self.doc_helper.converted_file})"
                        logger.debug(message)
                        Telegram.send_message(message)

    # sets the size and position of the window
    def set_windows_size_word(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2000, 1400, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    @staticmethod
    def prepare_document_for_test():
        KeyActions.click('/word/layout.png')
        sleep(0.3)
        KeyActions.click('/word/transfers.png')
        pg.press('down', interval=0.1)
        pg.press('enter')
        KeyActions.click('/powerpoint/view.png')
        sleep(0.3)
        KeyActions.click('/word/one_page.png')
        KeyActions.click('/word/resolution100.png')
        pg.moveTo(100, 0)
        sleep(0.5)

    def open_word_with_cmd(self, file_path):
        sb.Popen(f"{config.ms_office}/{ProjectConfig.WORD} -t {file_path}")
        self.waiting_for_opening_word()

    def check_open_word(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            match [win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)]:
                case ['OpusApp', window_text] if window_text != "":
                    self.windows_handler_number = hwnd
                case ['#32770', "Microsoft Word"]:
                    logger.debug(f"document recovery {self.doc_helper.converted_file}")
                    win32gui.SetForegroundWindow(hwnd)
                    pg.press('right')
                    pg.press('enter')

    def waiting_for_opening_word(self):
        self.windows_handler_number = None
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_word, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(config.wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open file {self.doc_helper.converted_file}")
                self.doc_helper.copy_testing_files_to_folder(self.doc_helper.too_long_to_open_files)
                break

    def errors_handler_when_opening(self):
        errors = self.get_errors()
        if errors:
            match errors:
                case ['#32770', 'Microsoft Word']:
                    logger.error(f"'an error has occurred while opening file: {self.doc_helper.converted_file}.")
                    pg.press('esc', presses=3, interval=0.2)
                    self.files_with_errors_when_opening.append(self.doc_helper.converted_file)
                    self.doc_helper.copy_testing_files_to_folder(self.doc_helper.opener_errors)
                    return False
                case _:
                    message = f"New Error\nModule: errors_handler_when_opening" \
                              f"Error message: {errors}\nFile: {self.doc_helper.converted_file}"
                    logger.debug(message)
                    Telegram.send_message(message)
                    self.doc_helper.copy_testing_files_to_folder(self.doc_helper.failed_source)
                    return False
        return True

    def events_handler_when_closing(self):
        errors = self.get_errors()
        if errors:
            match errors:
                case ["NUIDialog", "Microsoft Word"]:
                    print(f'Save file: {self.doc_helper.converted_file}')
                    pg.press('right')
                    pg.press('enter')
                case ["#32770", "Microsoft Word"]:
                    logger.debug(f"operation aborted {self.doc_helper.converted_file}")
                    pg.press('enter')
                case _:
                    message = f"New Error\nModule: events_handler_when_closing" \
                              f"Error message: {errors}\nFile: {self.doc_helper.converted_file}"
                    logger.debug(message)
                    Telegram.send_message(message)

    def events_handler_when_opening(self, opener_test=False):
        if opener_test:
            errors = self.get_errors()
            if errors:
                match errors:
                    case ['bosa_sdm_msword', 'Пароль']:
                        pg.press('tab')
                        pg.press('enter')
            return
        if self.get_errors():
            error_processing = Process(target=self.errors_handler_for_thread)
            error_processing.start()
            sleep(7)
            error_processing.terminate()

    def close_word_with_cmd(self):
        pg.hotkey('ctrl', 'z')
        pg.press('esc')
        FileUtils.run_command("taskkill /t /im  WINWORD.EXE")
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
                CompareImage.grab_coordinate(f"{path_to_save_screen}/page_{page_num}.png", coordinate)
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {self.doc_helper.converted_file}'
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
            Telegram.send_message(
                f'Exception while getting statistics, {self.doc_helper.converted_file}\nException: {e}')
            logger.exception(f'Exception while getting statistics, {self.doc_helper.converted_file}\nException: {e}')
            self.statistics_word = None

    def get_information_about_document(self, file_path):
        error_processing = Process(target=self.errors_handler_for_thread)
        error_processing.start()
        word_app = Dispatch('Word.Application')
        word_app.Visible = False
        try:
            word_app = word_app.Documents.Open(f'{file_path}', None, True)
            self.get_word_statistic(word_app)
            word_app.Close(False)
            self.num_of_page = self.statistics_word['num_of_sheets']
            return True
        except Exception as e:
            logger.exception(f"Can't get number of pages in {self.doc_helper.source_file}. Exception: {e}")
            self.doc_helper.copy_testing_files_to_folder(self.doc_helper.failed_source)
            return False
        finally:
            FileUtils.run_command("taskkill /t /im  WINWORD.EXE")
            error_processing.terminate()

    def statistic_report_generation(self, modified):
        modified_keys = [self.doc_helper.converted_file]
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
