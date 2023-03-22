# -*- coding: utf-8 -*-

import win32con
import win32gui
from loguru import logger
import pyautogui as pg

from frameworks.libre_office.macroses import Macroses


class LibreEvents:
    def __init__(self):
        self.events = []

    def enum_windows_callback(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if class_name in ['SALSUBFRAME', '#32770', 'SALFRAME']:
                if window_text in ['Восстановление документа LibreOffice 7.3', 'Отчёт о сбое', 'Сохранить документ?']:
                    win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                    win32gui.SetForegroundWindow(hwnd)
                    self.events = [class_name, window_text]

    def when_opening(self, file_name=None):
        self.events.clear()
        win32gui.EnumWindows(self.enum_windows_callback, self.events)
        if self.events and self.events[1] == "Восстановление документа LibreOffice 7.3":
            logger.debug(f"Восстановление документа LibreOffice 7.3 {file_name}")
            Macroses.close_file_recovery_window()
            self.events.clear()
        elif self.events and self.events[1] == 'Отчёт о сбое':
            logger.debug(f"Отчёт о сбое {file_name}")
            pg.press('esc', interval=0.5)
            self.events.clear()
            win32gui.EnumWindows(self.enum_windows_callback, self.events)
            if self.events and self.events[1] == "Восстановление документа LibreOffice 7.3":
                Macroses.close_file_recovery_window()
                self.events.clear()

    def when_closing(self):
        self.events.clear()
        win32gui.EnumWindows(self.enum_windows_callback, self.events)
        if self.events and self.events[1] == "Сохранить документ?":
            pg.press('right')
            pg.press('enter')
            self.events.clear()
