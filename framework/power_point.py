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

from config import wait_for_opening, wait_for_press
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


class PowerPoint:

    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.slide_count = None
        self.shell = Dispatch("WScript.Shell")
        self.click = self.helper.click
        self.errors_handler = False
        self.waiting_time = False

    def prepare_windows(self):
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

    # gets the coordinates of the window
    # sets the size and position of the window
    def get_coord_pp(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 0, 0, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def set_foreground_window(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(hwnd)
                pg.press('esc')

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' \
                    and win32gui.GetWindowText(hwnd) == "Microsoft PowerPoint" \
                    or win32gui.GetClassName(hwnd) == 'NUIDialog':
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.check_errors.errors.clear()
                self.check_errors.errors.append(win32gui.GetClassName(hwnd))
                self.check_errors.errors.append(win32gui.GetWindowText(hwnd))
                if win32gui.GetClassName(hwnd) == 'NUIDialog':
                    pg.press('enter')
                    sleep(2)
                    self.check_errors.errors.clear()

    def opener_power_point(self, path_for_open, file_name):
        error_processing = Process(target=self.check_errors.run_get_errors_pp, args=(self.helper.converted_file,))
        error_processing.start()
        presentation = Dispatch("PowerPoint.application")
        try:
            presentation = presentation.Presentations.Open(f'{path_for_open}{file_name}')
            self.slide_count = len(presentation.Slides)
            print(f"[bold blue]Number of Slides[/bold blue]:{self.slide_count}")
            return True

        except Exception:
            logger.error(f'Exception while opening presentation. {self.helper.converted_file}')
            self.slide_count = None
            self.helper.copy_to_folder(self.helper.failed_source)
            return False

        finally:
            error_processing.terminate()
            self.close_presentation(presentation)

    def close_presentation(self, presentation):
        try:
            presentation.close()
        except Exception:
            logger.debug(f'Exception while closing presentation. {self.helper.converted_file}')
        finally:
            os.system("taskkill /im  POWERPNT.EXE")

    def open_presentation_with_cmd(self, file_name):
        self.check_errors.errors.clear()
        self.helper.run(self.helper.tmp_dir_in_test, file_name, self.helper.power_point)
        self.waiting_for_opening_power_point()

    def check_open_power_point(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass' and win32gui.GetWindowText(hwnd) != '':
                self.waiting_time = True

    def waiting_for_opening_power_point(self):
        self.waiting_time = False
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_power_point, self.waiting_time)
            if self.waiting_time:
                sleep(2)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open "
                             f"Copied files: {self.helper.converted_file} "
                             f"and {self.helper.source_file} to 'failed_to_open_converted_file/too_long_to_open_files'")
                self.helper.copy_to_folder(self.helper.too_long_to_open_files)
                break

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == "#32770" \
                and self.check_errors.errors[1] == "Microsoft PowerPoint":

            logger.error(f"'an error has occurred while opening the file'. "
                         f"Copied files: {self.helper.converted_file} "
                         f"and {self.helper.source_file} to 'failed_to_open_converted_file'")

            pg.press('esc', presses=3, interval=0.2)
            self.helper.create_dir(self.helper.opener_errors)
            self.helper.copy_to_folder(self.helper.opener_errors)
            self.check_errors.errors.clear()
            return False

        elif not self.check_errors.errors:
            return True

        else:
            logger.debug(f"New Error "
                         f"Error message: {self.check_errors.errors} "
                         f"Filename: {self.helper.converted_file}")
            self.helper.copy_to_folder(self.helper.failed_source)
            return False

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == 'NUIDialog':
            pg.press('right')
            pg.press('enter')
            self.check_errors.errors.clear()

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshot(self, path_to_save_screen):
        if not self.check_errors.errors:
            win32gui.EnumWindows(self.get_coord_pp, self.coordinate)
            coordinate = self.coordinate[0]
            coordinate = (coordinate[0] + 350,
                          coordinate[1] + 170,
                          coordinate[2] - 120,
                          coordinate[3] - 100)

            self.prepare_windows()
            page_num = 1
            for page in range(self.slide_count):
                CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
                pg.press('pgdn')
                sleep(wait_for_press)
                page_num += 1

    def close_presentation_with_hotkey(self):
        win32gui.EnumWindows(self.set_foreground_window, self.coordinate)
        pg.hotkey('ctrl', 'z', interval=0.2)
        pg.hotkey('ctrl', 'q', interval=0.2)
        sleep(0.2)
        self.events_handler_when_closing()

