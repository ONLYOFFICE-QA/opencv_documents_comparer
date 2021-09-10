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

    def __init__(self):
        self.create_project_dirs()
        self.copy_for_test()
        self.coordinate = []
        # self.run_compare_pp(list_file_names_doc_from_compare,
        #                     extension_from,
        #                     extension_to)
        self.run_compare_pp(os.listdir(custom_doc_to),
                            extension_from,
                            extension_to)

    def get_coord_pp(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.MoveWindow(hwnd, 494, 30, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    @staticmethod
    def opener_power_point(path_for_open, file_name):
        try:
            presentation = Dispatch("PowerPoint.application")
            presentation = presentation.Presentations.Open(f'{path_for_open}{file_name}')
            slide_count = len(presentation.Slides)
            # print(presentation)
            print(slide_count)
            return slide_count, presentation
        except Exception:
            print('[bold red]NOT TESTED!!![/bold red]')
            sb.call(["taskkill", "/IM", "POWERPNT.EXE"])
            return 'error', 'error'

    def get_screenshot(self, file_name_for_screen, path_to_save_screen):
        print(f'[bold green]In test[/bold green] {file_name_for_screen}')
        slide_count, presentation = self.opener_power_point(path_to_folder_for_test, file_name_for_screen)
        win32gui.EnumWindows(self.get_coord_pp, self.coordinate)
        sleep(wait_for_open)
        page_num = 1
        if slide_count == 'error':
            return 'error'
        else:
            for page in range(slide_count):
                CompareImage.grab_coordinate(path_to_save_screen, file_name_for_screen, page_num, self.coordinate[0])
                pg.press('pgdn')
                sleep(wait_for_press)
                page_num += 1
            presentation.Close()
            return slide_count

    def run_compare_pp(self, list_of_files, from_extension, to_extension):
        for file_name in list_of_files:
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            if to_extension == file_name.split('.')[-1]:
                slide_count_after = self.get_screenshot(file_name,
                                                        path_to_tmpimg_after_conversion)
                if slide_count_after == 'error':
                    print('[bold red]ERROR slide_count_after[/bold red]')
                    self.copy_to_errors(file_name,
                                        file_name_from)
                else:
                    slide_count_before = self.get_screenshot(file_name_from,
                                                             path_to_tmpimg_befor_conversion)
                    if slide_count_before == 'error':
                        print('[bold red]ERROR slide_count_before[/bold red]')
                        self.copy_to_errors(file_name,
                                            file_name_from)

                    elif slide_count_after != slide_count_before:
                        self.copy_to_not_tested(file_name,
                                                file_name_from)
                        print('[bold red]SLIDE COUNT DIFFERENT[/bold red]')
                    else:
                        CompareImage()
        pass
