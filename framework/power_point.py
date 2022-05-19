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

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' \
                    and win32gui.GetWindowText(hwnd) == "Microsoft PptPptxCompareImg" \
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

        except Exception:
            logger.error(f'Exception while opening presentation. {self.helper.converted_file}')
            self.slide_count = None

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
        self.helper.run(self.helper.tmp_dir_in_test, file_name, self.helper.power_point)
        sleep(wait_for_opening)
        # check errors
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)

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

    def close_presentation_with_cmd(self):
        pg.hotkey('ctrl', 'z')
        os.system("taskkill /im  POWERPNT.EXE")
        sleep(0.2)
        win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == 'NUIDialog':
            pg.press('right')
            pg.press('enter')
            self.check_errors.errors.clear()
        pass

