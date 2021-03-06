# -*- coding: utf-8 -*-
import os
from time import sleep
from loguru import logger

import pyautogui as pg
import win32con
import win32gui
from win32com.client import Dispatch

from config import wait_for_opening, wait_for_press
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


class LibreOffice:

    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.shell = Dispatch("WScript.Shell")
        self.click = self.helper.click

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

    def get_coord(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass' \
                    or win32gui.GetClassName(hwnd) == 'SALFRAME':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 0, 0, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'SALSUBFRAME' \
                    or win32gui.GetClassName(hwnd) == '#32770' \
                    or win32gui.GetClassName(hwnd) == 'SALFRAME' \
                    and win32gui.GetWindowText(hwnd) == '???????????????????????????? ?????????????????? LibreOffice 7.3' \
                    or win32gui.GetWindowText(hwnd) == '?????????? ?? ????????':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.check_errors.errors.clear()
                self.check_errors.errors.append(win32gui.GetClassName(hwnd))
                self.check_errors.errors.append(win32gui.GetWindowText(hwnd))

    def open_libre_office_with_cmd(self, file_name):
        self.helper.run_libre_with_cmd(self.helper.tmp_dir_in_test, file_name)
        sleep(wait_for_opening)
        self.events_handler_when_opening()  # check events when opening

    def get_screenshot_odp(self, path_to_save_screen, slide_count):
        win32gui.EnumWindows(self.get_coord, self.coordinate)
        coordinate = self.coordinate[0]
        coordinate = (coordinate[0] + 350,
                      coordinate[1] + 170,
                      coordinate[2] - 120,
                      coordinate[3] - 100)

        self.prepare_windows_hot_keys()
        page_num = 1
        for page in range(slide_count):
            CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
            pg.press('pgdn')
            sleep(wait_for_press)
            page_num += 1

    def events_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[1] == "???????????????????????????? ?????????????????? LibreOffice 7.3":
            logger.debug(f"???????????????????????????? ?????????????????? LibreOffice 7.3 {self.helper.converted_file}")
            self.close_file_recovery_window()
            self.check_errors.errors.clear()
        elif self.check_errors.errors \
                and self.check_errors.errors[1] == '?????????? ?? ????????':
            logger.debug(f"?????????? ?? ???????? {self.helper.converted_file}")
            pg.press('esc', interval=0.5)
            self.check_errors.errors.clear()
            win32gui.EnumWindows(self.check_error, self.check_errors.errors)
            if self.check_errors.errors \
                    and self.check_errors.errors[1] == "???????????????????????????? ?????????????????? LibreOffice 7.3":
                self.close_file_recovery_window()
                self.check_errors.errors.clear()

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors and self.check_errors.errors[1] == "?????????????????? ?????????????????":
            pg.press('right')
            pg.press('enter')
            self.check_errors.errors.clear()

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[1] == "????????????":
            logger.error(f"'an error has occurred while opening the file'. "
                         f"Copied files: {self.helper.converted_file} "
                         f"and {self.helper.source_file} to 'failed_to_open_converted_file'")
            self.helper.create_dir(self.helper.opener_errors)
            self.helper.copy_to_folder(self.helper.opener_errors)
            pg.press('enter')
            self.check_errors.errors.clear()
        elif self.check_errors.errors:
            logger.debug(f"Error message: {self.check_errors.errors} "
                         f"Filename: {self.helper.converted_file}")
            self.check_errors.errors.clear()

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
        os.system("taskkill /im  soffice.bin")
