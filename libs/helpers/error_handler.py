# -*- coding: utf-8 -*-
import subprocess as sb
from time import sleep

import pyautogui as pg
import win32gui
from loguru import logger


class CheckErrors:
    def __init__(self):
        self.errors = []

    def get_windows_title(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' \
                    or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword' \
                    or win32gui.GetClassName(hwnd) == 'ThunderDFrame' \
                    or win32gui.GetClassName(hwnd) == 'NUIDialog':
                win32gui.SetForegroundWindow(hwnd)

                self.errors.clear()
                self.errors.append(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetWindowText(hwnd))

    def run_get_error_exel(self, filename):
        while True:
            win32gui.EnumWindows(self.get_windows_title, self.errors)
            if self.errors:
                self.check_errors_excel(self.errors, filename)
                self.errors.clear()

    def run_get_errors_word(self, filename):
        while True:
            win32gui.EnumWindows(self.get_windows_title, self.errors)
            if self.errors:
                self.check_word(self.errors, filename)
                self.errors.clear()

    def run_get_errors_pp(self, filename):
        while True:
            win32gui.EnumWindows(self.get_windows_title, self.errors)
            if self.errors:
                self.check_pp(self.errors, filename)
                self.errors.clear()

    @staticmethod
    def check_errors_excel(array_of_errors, filename):
        logger.error(f'"{array_of_errors}" happened while opening: {filename}')
        if array_of_errors[0] == '#32770':
            if array_of_errors[1] == 'Microsoft Visual Basic':
                try:
                    pg.click("libs/image_templates/excel/end.png")
                except Exception:
                    sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)

            elif array_of_errors[1] == '???????????????? ???????????????????? ??????????????????':
                pass

            elif array_of_errors[1] == 'Microsoft Excel':
                pg.press('enter')
                sleep(1)
                pg.press('enter')

            elif array_of_errors[1] == 'Monopoly':
                pg.press('enter')

        elif array_of_errors[0] == 'ThunderDFrame':
            if array_of_errors[1] == 'Functions List':
                pg.hotkey('alt', 'f4')

            elif array_of_errors[1] == 'Select Players and Times':
                pg.press('tab', presses=6, interval=0.2)
                pg.press('enter', interval=0.2)

        elif array_of_errors[0] == 'NUIDialog':
            if array_of_errors[1] == 'Microsoft Excel' \
                    or array_of_errors[1] == 'Microsoft Excel - ???????????????? ??????????????????????????':
                pg.press('enter')
        else:
            logger.debug(f'"New Event {array_of_errors}" happened while opening: {filename}')

    @staticmethod
    def check_word(array_of_errors, filename):
        if not array_of_errors[1] == '???????????????????????????? ??????????':
            logger.error(f'"{array_of_errors}" happened while opening: {filename}')
        if array_of_errors[0] == '#32770':
            if array_of_errors[1] == 'Microsoft Word':
                pg.press('left')
                pg.press('enter')

            elif array_of_errors[1] == 'Microsoft Visual Basic for Applications':
                pg.press('enter')

            elif array_of_errors[1] == '???????????????? ???????????????????? ??????????????????':
                print(array_of_errors[1])

            elif array_of_errors[1] == '?????????????????? ??????':
                sb.call(f'powershell.exe kill -Name WINWORD', shell=True)

        elif array_of_errors[0] == 'bosa_sdm_msword':
            if array_of_errors[1] == '???????????????????????????? ??????????':
                pg.press('enter')

            elif array_of_errors[1] == '????????????':
                pg.press('tab')
                pg.press('enter')

            elif array_of_errors[1] == '???????????????? ??????????????????????':
                sleep(2)
                pg.press('tab', presses=3)
                pg.press('enter')

        elif array_of_errors[0] == 'NUIDialog' and array_of_errors[1] == 'Microsoft Word':
            pg.press('enter')

        elif array_of_errors[0] == 'NUIDialog' \
                and array_of_errors[1] == '?????????????????? ?????????????? ???????????????????????? Microsoft Word':
            pg.press('left')
            pg.press('enter')

        else:
            logger.debug(f'"New Event {array_of_errors}" happened while opening: {filename}')

    @staticmethod
    def check_pp(array_of_errors, filename):
        logger.error(f'"{array_of_errors}" happened while opening: {filename}')
        if array_of_errors[0] == 'NUIDialog':
            if array_of_errors[1] == '????????????':
                pg.press('right', presses=2)
                pg.press('enter')
        else:
            logger.debug(f'"New Event {array_of_errors}" happened while opening: {filename}')
