# -*- coding: utf-8 -*-
from time import sleep
from loguru import logger

import pyautogui as pg
import win32con
import win32gui

from config import wait_for_opening
from framework.telegram import Telegram
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


class LibreOffice:

    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.errors = self.check_errors.errors
        self.click = self.helper.click
        self.windows_handler_number = None
        self.errors_files_when_opening = []

    @staticmethod
    def prepare_windows_hot_keys():
        pg.press('alt', interval=0.1)
        pg.press('right', presses=2, interval=0.1)
        pg.press('down', interval=0.1)
        pg.press('enter', interval=0.1)
        pg.press('alt', interval=0.1)
        pg.press('right', presses=2, interval=0.1)
        pg.press('up', presses=2, interval=0.1)
        pg.press('enter', interval=0.1)
        pg.press('up', interval=0.1)
        pg.press('enter', interval=0.1)
        pg.hotkey('ctrl', 'a', interval=0.1)
        sleep(0.1)
        pg.write('100', interval=0.1)
        pg.press('enter', interval=0.1)
        sleep(0.5)

    def set_windows_size_word(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2200, 1420, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            window_text_errors_array = ['Восстановление документа LibreOffice 7.3', 'Отчёт о сбое', 'Ошибка']
            if class_name == 'SALSUBFRAME' or class_name == '#32770' or class_name == 'SALFRAME':
                if window_text in window_text_errors_array:
                    win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                    win32gui.SetForegroundWindow(hwnd)
                    self.errors.clear()
                    self.errors.append(win32gui.GetClassName(hwnd))
                    self.errors.append(win32gui.GetWindowText(hwnd))

    def open_libre_office_with_cmd(self, file_name):
        self.errors.clear()
        self.helper.run_libre_with_cmd(self.helper.tmp_dir_in_test, file_name)
        self.waiting_for_opening_libre_office()
        self.events_handler_when_opening()  # check events when opening

    def check_open_libre_office(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'SALFRAME' and win32gui.GetWindowText(hwnd) != '' \
                    or win32gui.GetClassName(hwnd) == 'SALSUBFRAME':
                self.windows_handler_number = hwnd

    def waiting_for_opening_libre_office(self):
        self.windows_handler_number = None
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_libre_office, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open file: {self.helper.converted_file}")
                self.helper.copy_to_folder(self.helper.opener_errors)
                break

    def get_coordinate_libreoffice(self):
        coordinate = [win32gui.GetWindowRect(self.windows_handler_number)]
        coordinate = coordinate[0]
        coordinate = (coordinate[0] + 350,
                      coordinate[1] + 170,
                      coordinate[2] - 120,
                      coordinate[3] - 100)
        return coordinate

    def get_screenshot_odp(self, path_to_save_screen, slide_count):
        if win32gui.IsWindow(self.windows_handler_number):
            self.set_windows_size_word()
            coordinate = self.get_coordinate_libreoffice()
            self.prepare_windows_hot_keys()
            page_num = 1
            for page in range(slide_count):
                CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {self.helper.converted_file}'
            Telegram.send_message(massage)
            logger.error(massage)

    def events_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.errors)
        if self.errors and self.errors[1] == "Восстановление документа LibreOffice 7.3":
            logger.debug(f"Восстановление документа LibreOffice 7.3 {self.helper.converted_file}")
            self.close_file_recovery_window()
            self.errors.clear()
        elif self.errors and self.errors[1] == 'Отчёт о сбое':
            logger.debug(f"Отчёт о сбое {self.helper.converted_file}")
            pg.press('esc', interval=0.5)
            self.errors.clear()
            win32gui.EnumWindows(self.check_error, self.errors)
            if self.errors and self.errors[1] == "Восстановление документа LibreOffice 7.3":
                self.close_file_recovery_window()
                self.errors.clear()

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.check_error, self.errors)
        if self.errors and self.errors[1] == "Сохранить документ?":
            pg.press('right')
            pg.press('enter')
            self.errors.clear()

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.errors)
        if self.errors and self.errors[1] == "Ошибка":
            logger.error(f"'an error has occurred while opening the file' File: {self.helper.converted_file}")
            self.errors_files_when_opening.append(self.helper.converted_file)
            self.helper.copy_to_folder(self.helper.opener_errors)
            pg.press('enter')
            self.errors.clear()
        elif self.errors:
            logger.debug(f"Error message: {self.errors} Filename: {self.helper.converted_file}")
            self.errors.clear()

    @staticmethod
    def close_file_recovery_window():
        pg.press('esc', interval=0.2)
        pg.press('left', interval=0.2)
        pg.press('enter', interval=0.2)
        sleep(wait_for_opening)

    def close_libre(self):
        pg.hotkey('ctrl', 'q')
        sleep(0.5)
        self.events_handler_when_closing()  # check events when closing
