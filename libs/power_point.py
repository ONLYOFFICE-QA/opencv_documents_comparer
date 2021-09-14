import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print
from win32com.client import Dispatch

from libs.Compare_Image import CompareImage
from libs.helper import Helper
from var import *


class PowerPoint(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        self.copy_for_test(list_of_files)
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
            # print(presentation)
            print(slide_count)
            presentation.Close()
            # sleep(5)
            # sb.call("powershell.exe kill -Name POWERPNT", shell=True)
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return slide_count

        except Exception:
            print('[bold red]NOT TESTED!!![/bold red]')
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])

    def get_screenshot(self, path_to_save_screen, file_name, slide_count):
        print(f'[bold green]In test[/bold green] {file_name}')
        self.run(path_to_folder_for_test, file_name, "POWERPNT.EXE")
        sleep(wait_for_open)
        win32gui.EnumWindows(self.check_error, self.errors)
        print(f'Step 1 {self.errors}')
        if not self.errors:
            print(f'error = {self.errors}')
            win32gui.EnumWindows(self.get_coord_pp, self.coordinate)
            print(f'Step 2 {self.coordinate}')
            page_num = 1
            for page in range(slide_count):
                CompareImage.grab_coordinate_pp(path_to_save_screen, file_name, page_num, self.coordinate[0])
                pg.press('pgdn')
                sleep(wait_for_press)
                page_num += 1
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return 'normal'
        elif self.errors and self.errors[0] == '#32770':
            print(f'Step 3 {self.errors[0]}')
            pg.press('enter')
            sb.call(["TASKKILL", "/IM", "POWERPNT.EXE", "/t", "/f"], shell=True)
            return self.errors[0]
        sb.call(["taskkill", "/IM", "POWERPNT.EXE"])

    def run_compare_pp(self, list_of_files):
        for file_name in list_of_files:
            file_name_from = file_name.replace(f'.{extension_to}', f'.{extension_from}')
            if extension_to == file_name.split('.')[-1]:
                self.copy(path_to_folder_for_test + file_name_from, path_to_temp_in_test + file_name_from)
                slide_count = self.opener_power_point(path_to_temp_in_test,
                                                      file_name_from)
                self.delete(path_to_temp_in_test + file_name_from)
                print(slide_count)

                error = self.get_screenshot(tmp_after,
                                            file_name,
                                            slide_count)
                print(error)
                if error == "#32770":
                    print('[bold red]ERROR, copy to not tested folder[/bold red]')
                    self.copy_to_not_tested(file_name, file_name_from)
                    self.errors.clear()

                elif error != "#32770":
                    self.get_screenshot(tmp_befor,
                                        file_name_from,
                                        slide_count)
                    sb.call(["TASKKILL", "/IM", "POWERPNT.EXE", "/t", "/f"], shell=True)
                    CompareImage(file_name)

                pass
            pass
        pass

    pass
