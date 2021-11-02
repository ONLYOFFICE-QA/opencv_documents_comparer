# -*- coding: utf-8 -*-
import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print
from win32com.client import Dispatch

from config import *
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper

source_extension = 'ppt'
converted_extension = 'pptx'


class PowerPoint:

    def __init__(self):
        self.helper = Helper(source_extension, converted_extension)
        self.coordinate = []
        self.errors = []
        self.click = self.helper.click

    def prepare_windows(self):
        self.click('libs/image_templates/excel/turn_on_content.png')
        self.click('libs/image_templates/excel/turn_on_content.png')
        sleep(0.2)
        self.click('libs/image_templates/powerpoint/ok.png')
        self.click('libs/image_templates/powerpoint/view.png')
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
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 0, 0, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    # Checks the window title
    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.errors.clear()
                print(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetClassName(hwnd))
            elif win32gui.GetClassName(hwnd) == 'NUIDialog':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.errors.clear()
                print(win32gui.GetClassName(hwnd))
                pg.press('enter')
                sleep(2)

    @staticmethod
    def opener_power_point(path_for_open, file_name):
        try:
            presentation = Dispatch("PowerPoint.application")
            presentation = presentation.Presentations.Open(f'{path_for_open}{file_name}')
            slide_count = len(presentation.Slides)
            print(slide_count)
            presentation.Close()
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return slide_count

        except Exception:
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return 'None'

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshot(self, path_to_save_screen, file_name, slide_count):
        self.helper.run(self.helper.tmp_dir_in_test, file_name, self.helper.power_point)
        sleep(wait_for_opening)
        win32gui.EnumWindows(self.check_error, self.errors)
        if not self.errors:
            win32gui.EnumWindows(self.get_coord_pp, self.coordinate)
            coordinate = self.coordinate[0]
            coordinate = (coordinate[0] + 350,
                          coordinate[1] + 170,
                          coordinate[2] - 120,
                          coordinate[3] - 100)
            self.prepare_windows()
            page_num = 1
            for page in range(slide_count):
                CompareImage.grab_coordinate(path_to_save_screen, file_name, page_num, coordinate)
                pg.press('pgdn')
                sleep(wait_for_press)
                page_num += 1
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return 'successes'
        elif self.errors and self.errors[0] == '#32770':
            pg.press('enter')
            sb.call(["TASKKILL", "/IM", "POWERPNT.EXE", "/t", "/f"], shell=True)
            return self.errors[0]

    def run_compare_pp(self, list_of_files):
        for converted_file in list_of_files:
            if converted_file.endswith((".pptx", ".PPTX")):
                source_file, tmp_name_converted_file, \
                tmp_name_source_file, tmp_name = self.helper.preparing_files_for_test(converted_file,
                                                                                      converted_extension,
                                                                                      source_extension)
                slide_count = self.opener_power_point(self.helper.tmp_dir_in_test,
                                                      tmp_name)

                if slide_count == 'None':
                    print("[bold red]Can't open source file[/bold red]")
                    self.helper.copy_to_folder(converted_file,
                                               source_file,
                                               self.helper.failed_source)

                else:
                    print(f'[bold green]In test[/bold green] {converted_file}')
                    error = self.get_screenshot(self.helper.tmp_dir_converted_image,
                                                tmp_name_converted_file,
                                                slide_count)

                    if error == "#32770":
                        print('[bold red]ERROR, copy to "not tested" [/bold red]')
                        self.helper.copy_to_folder(converted_file,
                                                   source_file,
                                                   self.helper.untested_folder)
                        self.errors.clear()

                    else:
                        self.get_screenshot(self.helper.tmp_dir_source_image,
                                            tmp_name_source_file,
                                            slide_count)

                        CompareImage(converted_file, self.helper)

        self.helper.delete(f'{self.helper.tmp_dir_in_test}*')
        pass

    pass
