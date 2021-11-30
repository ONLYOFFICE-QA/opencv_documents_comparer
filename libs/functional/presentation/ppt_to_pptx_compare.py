# -*- coding: utf-8 -*-
import os
import subprocess as sb
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from rich import print
from win32com.client import Dispatch

from config import *
from libs.helpers.compare_image import CompareImage
from libs.helpers.get_error import CheckErrors
from libs.helpers.helper import Helper

source_extension = 'ppt'
converted_extension = 'pptx'


class PowerPoint:

    def __init__(self):
        self.check_errors = CheckErrors()
        self.helper = Helper(source_extension, converted_extension)
        self.coordinate = []
        self.shell = Dispatch("WScript.Shell")
        self.click = self.helper.click
        logger.add('./logs/ppt_pptx.log',
                   format="{time} {level} {message}",
                   level="DEBUG",
                   mode='w')

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

    def opener_power_point(self, path_for_open, file_name, file_name_for_log):
        error_processing = Process(target=self.check_errors.run_get_errors_pp, args=(file_name_for_log,))
        error_processing.start()
        try:
            presentation = Dispatch("PowerPoint.application")
            presentation = presentation.Presentations.Open(f'{path_for_open}{file_name}')
            slide_count = len(presentation.Slides)
            print(f"[bold blue]Number of Slides[/bold blue]:{slide_count}")
            presentation.Close()
            return slide_count

        except Exception:
            return 'None'

        finally:
            error_processing.terminate()
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"], shell=True)

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshot(self, path_to_save_screen, file_name, slide_count):
        self.helper.run(self.helper.tmp_dir_in_test, file_name, self.helper.power_point)
        sleep(wait_for_opening)
        # check errors
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if not self.check_errors.errors:
            win32gui.EnumWindows(self.get_coord_pp, self.coordinate)
            coordinate = self.coordinate[0]
            coordinate = (coordinate[0] + 350,
                          coordinate[1] + 170,
                          coordinate[2] - 120,
                          coordinate[3] - 100)

            self.prepare_windows()
            page_num = 1
            for page in range(slide_count):
                CompareImage.grab_coordinate(path_to_save_screen, page_num, coordinate)
                pg.press('pgdn')
                sleep(wait_for_press)
                page_num += 1
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])

    def run_compare_pp(self, list_of_files):
        for converted_file in list_of_files:
            if converted_file.endswith((".pptx", ".PPTX")):
                source_file, tmp_name_converted_file, \
                tmp_name_source_file, tmp_name = self.helper.preparing_files_for_test(converted_file,
                                                                                      converted_extension,
                                                                                      source_extension)

                slide_count = self.opener_power_point(self.helper.tmp_dir_in_test, tmp_name, source_file)

                if slide_count != 'None':
                    print(f'[bold green]In test[/bold green] {converted_file}')
                    self.get_screenshot(self.helper.tmp_dir_converted_image,
                                        tmp_name_converted_file,
                                        slide_count)

                    if self.check_errors.errors \
                            and self.check_errors.errors[0] == "#32770" \
                            and self.check_errors.errors[1] == "Microsoft PowerPoint":
                        logger.error(f"'an error has occurred while opening the file'. "
                                     f"Copied files: {converted_file} and {source_file} to 'untested'")
                        pg.press('enter')
                        os.system("taskkill /t /f /im  POWERPNT.EXE")
                        self.helper.copy_to_folder(converted_file,
                                                   source_file,
                                                   self.helper.untested_folder)
                        self.check_errors.errors.clear()
                    elif not self.check_errors.errors:
                        print(f'[bold green]In test[/bold green] {source_file}')
                        self.get_screenshot(self.helper.tmp_dir_source_image,
                                            tmp_name_source_file,
                                            slide_count)

                        CompareImage(converted_file, self.helper)
                    else:
                        logger.debug(f"Error message: {self.check_errors.errors} Filename: {converted_file}")

                else:
                    logger.error(f"Can't open {source_file}. Copied files"
                                 "to 'failed_to_open_file'")

                    self.helper.copy_to_folder(converted_file,
                                               source_file,
                                               self.helper.failed_source)

            self.helper.tmp_cleaner()
