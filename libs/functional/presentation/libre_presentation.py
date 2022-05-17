# -*- coding: utf-8 -*-

from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from win32com.client import Dispatch

from config import *
from libs.helpers.compare_image import CompareImage
from libs.helpers.get_error import CheckErrors


class LibrePresentation:

    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.shell = Dispatch("WScript.Shell")
        self.click = self.helper.click
        logger.info(f'The {self.helper.source_extension}_{self.helper.converted_extension}'
                    f'comparison on version: {version} is running.')

    @staticmethod
    def prepare_windows_hot_keys():
        pg.press('alt')
        pg.press('right', presses=2)
        pg.press('down')
        pg.press('enter')

        pg.press('alt')
        pg.press('right', presses=2)
        pg.press('up', presses=2)
        pg.press('enter')
        pg.press('up')
        pg.press('enter')
        pg.hotkey('ctrl', 'a')
        sleep(0.1)
        pg.write('100', interval=0.1)
        pg.press('enter')
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
            if win32gui.GetClassName(hwnd) == 'SALFRAME':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.check_errors.errors.clear()
                self.check_errors.errors.append(win32gui.GetClassName(hwnd))
                self.check_errors.errors.append(win32gui.GetWindowText(hwnd))

    def open_presentation_with_cmd(self, file_name):
        self.helper.run_libre_with_cmd(self.helper.tmp_dir_in_test, file_name)
        sleep(wait_for_opening)
        # check errors
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)

    def get_screenshot_odp(self, path_to_save_screen, slide_count):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors:
            if self.check_errors.errors[0] == 'SALFRAME' and self.check_errors.errors[1] == 'Восстановление документа LibreOffice 7.2':
                pg.press('tab', interval=0.05)
                pg.press('right', interval=0.05)
                pg.press('enter', interval=0.05)
                pg.press('left', interval=0.05)
                pg.press('enter', interval=0.05)
                sleep(1)

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

    @staticmethod
    def close_presentation():
        pg.hotkey('ctrl', 'q')
