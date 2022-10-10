# -*- coding: utf-8 -*-
import os
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from rich import print
from win32com.client import Dispatch

from config import wait_for_opening
from framework.telegram import Telegram
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


class PowerPoint:

    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.errors = self.check_errors.errors
        self.click = self.helper.click
        self.slide_count = None
        self.windows_handler_number = None
        self.errors_files_when_opening = []

    def prepare_presentation_for_test(self):
        self.click('libs/image_templates/excel/turn_on_content.png')
        self.click('libs/image_templates/excel/turn_on_content.png')
        sleep(0.2)
        self.click('libs/image_templates/powerpoint/ok.png')
        self.click('libs/image_templates/powerpoint/view.png')
        pg.click()
        sleep(0.2)
        self.click('libs/image_templates/powerpoint/normal_view.png')
        self.click('libs/image_templates/powerpoint/scale.png')
        pg.press('tab')
        pg.write('100', interval=0.1)
        pg.press('enter')
        sleep(0.5)

    # sets the size and position of the window
    def set_windows_size_pp(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2200, 1420, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if class_name == '#32770' and window_text == "Microsoft PowerPoint" or class_name == 'NUIDialog':
                win32gui.SetForegroundWindow(hwnd)
                self.errors.clear()
                self.errors.append(class_name)
                self.errors.append(window_text)
                if class_name == 'NUIDialog':
                    pg.press('enter')
                    sleep(2)
                    self.errors.clear()

    def opener_power_point(self, path_for_open, file_name):
        error_processing = Process(target=self.check_errors.run_get_errors_pp, args=(self.helper.converted_file,))
        error_processing.start()
        presentation = Dispatch("PowerPoint.application")
        try:
            presentation = presentation.Presentations.Open(f'{path_for_open}{file_name}')
            self.slide_count = len(presentation.Slides)
            print(f"[bold blue]Number of Slides[/bold blue]:{self.slide_count}")
            return True

        except Exception as e:
            logger.error(f'Exception while opening presentation. {self.helper.converted_file}\nException: {e}')
            self.slide_count = None
            self.helper.copy_to_folder(self.helper.failed_source)
            return False

        finally:
            error_processing.terminate()
            self.close_presentation_with_hotkey()
            os.system("taskkill /im  POWERPNT.EXE")

    def open_presentation_with_cmd(self, file_name):
        self.errors.clear()
        self.helper.run(self.helper.tmp_dir_in_test, file_name, self.helper.power_point)
        self.waiting_for_opening_power_point()

    def check_open_power_point(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass' and win32gui.GetWindowText(hwnd) != '':
                self.windows_handler_number = hwnd

    def waiting_for_opening_power_point(self):
        self.windows_handler_number = None
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_power_point, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open file: {self.helper.converted_file} ")
                self.helper.copy_to_folder(self.helper.too_long_to_open_files)
                break

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.errors)
        if self.errors and self.errors[0] == "#32770" and self.errors[1] == "Microsoft PowerPoint":
            logger.error(f"'an error has occurred while opening the file' Files: {self.helper.converted_file}")
            self.errors_files_when_opening.append(self.helper.converted_file)
            pg.press('esc', presses=3, interval=0.2)
            self.helper.copy_to_folder(self.helper.opener_errors)
            self.errors.clear()
            return False
        elif not self.errors:
            return True
        else:
            logger.debug(f"New Error\nError message: {self.errors}\nFile: {self.helper.converted_file}")
            self.helper.copy_to_folder(self.helper.failed_source)
            return False

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.errors)
        if self.errors and self.errors[0] == 'NUIDialog':
            pg.press('right')
            pg.press('enter')
            self.errors.clear()

    # gets the coordinates of the window
    def get_coordinate_pp(self):
        coordinate = [win32gui.GetWindowRect(self.windows_handler_number)]
        coordinate = coordinate[0]
        coordinate = (coordinate[0] + 350,
                      coordinate[1] + 170,
                      coordinate[2] - 120,
                      coordinate[3] - 100)
        return coordinate

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshot(self, path_to_save_screen):
        if win32gui.IsWindow(self.windows_handler_number):
            self.set_windows_size_pp()
            coordinate = self.get_coordinate_pp()
            self.prepare_presentation_for_test()
            page_num = 1
            for page in range(self.slide_count):
                CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {self.helper.converted_file}'
            Telegram.send_message(massage)
            logger.error(massage)

    def close_presentation_with_hotkey(self):
        if win32gui.IsWindow(self.windows_handler_number):
            win32gui.SetForegroundWindow(self.windows_handler_number)
        pg.hotkey('ctrl', 'z', interval=0.2)
        pg.hotkey('ctrl', 'q', interval=0.2)
        sleep(0.2)
        self.events_handler_when_closing()
