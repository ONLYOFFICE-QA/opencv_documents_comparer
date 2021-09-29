import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print
from win32com.client import Dispatch

from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper
from var import *

extension_source = 'ppt'
extension_converted = 'pptx'


class PowerPoint(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        self.delete(f'{tmp_in_test}*')
        self.delete(f'{tmp_converted_image}*')
        self.delete(f'{tmp_source_image}*')
        self.coordinate = []
        self.errors = []
        self.run_compare_pp(list_of_files)

    def get_coord_pp(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 494, 30, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.errors.clear()
                print(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetClassName(hwnd))

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
            print('[bold red]NOT TESTED!!![/bold red]')
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return 'None'

    def get_screenshot(self, path_to_save_screen, file_name, slide_count):
        self.run(tmp_in_test, file_name, power_point)
        sleep(wait_for_opening)
        win32gui.EnumWindows(self.check_error, self.errors)
        if not self.errors:
            win32gui.EnumWindows(self.get_coord_pp, self.coordinate)
            coordinate = self.coordinate[0]
            coordinate = (coordinate[0] + 350,
                          coordinate[1] + 170,
                          coordinate[2] - 50,
                          coordinate[3] - 20)
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
        sb.call(["TASKKILL", "/IM", "POWERPNT.EXE", "/t", "/f"], shell=True)

    def run_compare_pp(self, list_of_files):
        for converted_file in list_of_files:
            if converted_file.endswith((".pptx", ".PPTX")):
                source_file, tmp_name_converted_file, \
                tmp_name_source_file, tmp_name = self.preparing_files_for_test(converted_file,
                                                                               extension_converted,
                                                                               extension_source)
                slide_count = self.opener_power_point(tmp_in_test,
                                                      tmp_name)

                if slide_count == 'None':
                    print("[bold red]Can't open source file[/bold red]")
                    self.copy_to_folder(converted_file, source_file, failed_source)

                else:
                    print(f'[bold green]In test[/bold green] {converted_file}')
                    error = self.get_screenshot(tmp_converted_image,
                                                tmp_name_converted_file,
                                                slide_count)
                    print(error)
                    if error == "#32770":
                        print('[bold red]ERROR, copy to "not tested" [/bold red]')
                        self.copy_to_folder(converted_file,
                                            source_file,
                                            untested_folder)
                        self.errors.clear()

                    elif error != "#32770":
                        self.get_screenshot(tmp_source_image,
                                            tmp_name_source_file,
                                            slide_count)
                        sb.call(["TASKKILL", "/IM", "POWERPNT.EXE", "/t", "/f"], shell=True)
                        CompareImage(converted_file)

        self.delete(tmp_in_test)
        pass

    pass
