# -*- coding: utf-8 -*-

import pyautogui as pg
import win32con
import win32gui
from loguru import logger


class LibreErrors:
    def __init__(self):
        self.errors: list = []
        self.errors_files_when_opening = []

    def check_error(self, hwnd, windows_number):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            window_text_errors = ['Восстановление документа LibreOffice 7.3', 'Отчёт о сбое', 'Ошибка']
            if class_name == 'SALSUBFRAME' or class_name == '#32770' or class_name == 'SALFRAME':
                if window_text in window_text_errors:
                    win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                    windows_number.append(hwnd)
                    self.errors = [win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)]

    def when_opening(self, file_name: str = None):
        windows_number = []
        win32gui.EnumWindows(self.check_error, windows_number)
        if self.errors and self.errors[1] == "Ошибка":
            logger.error(f"\n'an error has occurred while opening the file' File: {file_name}")
            self.errors_files_when_opening.append(file_name)
            self.errors.clear()
        elif self.errors:
            logger.debug(f"Error message: {self.errors} Filename: {file_name}")
            self.errors.clear()
