# -*- coding: utf-8 -*-
import subprocess as sb
from os.path import join
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger

import settings as config
from frameworks.StaticData import StaticData
from frameworks.host_control import FileUtils
from frameworks.image_handler import CompareImage
from frameworks.telegram import Telegram
from .macroses import Macroses


class LibreOffice:

    def __init__(self):
        self.doc_helper = StaticData.DOC_ACTIONS
        self.errors = []
        self.windows_handler_number = None
        self.errors_files_when_opening = []
        FileUtils.create_dir(StaticData.TMP_DIR_CONVERTED_IMG)
        FileUtils.create_dir(StaticData.TMP_DIR_SOURCE_IMG)

    def set_windows_size_word(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2200, 1420, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            window_text_errors = ['Восстановление документа LibreOffice 7.3', 'Отчёт о сбое', 'Ошибка']
            if class_name == 'SALSUBFRAME' or class_name == '#32770' or class_name == 'SALFRAME':
                if window_text in window_text_errors:
                    win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                    win32gui.SetForegroundWindow(hwnd)
                    self.errors.clear()
                    self.errors.append(win32gui.GetClassName(hwnd))
                    self.errors.append(win32gui.GetWindowText(hwnd))

    def open_file(self, file_path):
        self.errors.clear()
        sb.Popen(f"{join(config.libre_office, StaticData.LIBRE)} -o {file_path}")
        self.wait_until_open()
        self.events_handler_when_opening()  # check events when opening

    def wait_until_open(self):
        self.windows_handler_number = None
        stop_waiting = 1

        def wrapper(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                if (win32gui.GetClassName(hwnd) == 'SALFRAME' and win32gui.GetWindowText(hwnd) != '') or \
                        (win32gui.GetClassName(hwnd) == 'SALSUBFRAME'):
                    self.windows_handler_number = hwnd

        while True:
            win32gui.EnumWindows(wrapper, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(config.wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open file: {self.doc_helper.converted_file}")
                self.doc_helper.copy_testing_files_to_folder(self.doc_helper.opener_errors)
                break

    def get_coordinate(self):
        coordinate = [win32gui.GetWindowRect(self.windows_handler_number)][0]
        return (
            coordinate[0] + 370,
            coordinate[1] + 200,
            coordinate[2] - 120,
            coordinate[3] - 100
        )

    def get_screenshot_odp(self, path_to_save_screen, slide_count):
        if win32gui.IsWindow(self.windows_handler_number):
            self.set_windows_size_word()
            Macroses.prepare_windows_hot_keys()
            page_num = 1
            for page in range(slide_count):
                CompareImage.grab_coordinate(join(path_to_save_screen, f'page_{page_num}.png'), self.get_coordinate())
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {self.doc_helper.converted_file}'
            Telegram.send_message(massage)
            logger.error(massage)

    def events_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.errors)
        if self.errors and self.errors[1] == "Восстановление документа LibreOffice 7.3":
            logger.debug(f"Восстановление документа LibreOffice 7.3 {self.doc_helper.converted_file}")
            Macroses.close_file_recovery_window()
            self.errors.clear()
        elif self.errors and self.errors[1] == 'Отчёт о сбое':
            logger.debug(f"Отчёт о сбое {self.doc_helper.converted_file}")
            pg.press('esc', interval=0.5)
            self.errors.clear()
            win32gui.EnumWindows(self.check_error, self.errors)
            if self.errors and self.errors[1] == "Восстановление документа LibreOffice 7.3":
                Macroses.close_file_recovery_window()
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
            logger.error(f"'an error has occurred while opening the file' File: {self.doc_helper.converted_file}")
            self.errors_files_when_opening.append(self.doc_helper.converted_file)
            self.doc_helper.copy_testing_files_to_folder(self.doc_helper.opener_errors)
            pg.press('enter')
            self.errors.clear()
        elif self.errors:
            logger.debug(f"Error message: {self.errors} Filename: {self.doc_helper.converted_file}")
            self.errors.clear()

    def close_libre(self):
        pg.hotkey('ctrl', 'q')
        sleep(0.5)
        self.events_handler_when_closing()  # check events when closing
