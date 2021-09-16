import csv
import io
import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print

from libs.Compare_Image import CompareImage
from libs.Exel import Exel
from libs.helper import Helper
from var import *


class ExelCompareImage(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        # self.copy_for_test(list_of_files)
        self.coordinate = []
        self.errors = []
        self.run_compare_exel(list_of_files)

    def get_coord_exel(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'XLMAIN':
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                # win32gui.MoveWindow(hwnd, 494, 30, 2200, 1420, True)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def get_windows_title(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword':
                # hwnd = win32gui.FindWindow(None, "Telegram (15125)")
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)

                self.errors.clear()
                self.errors.append(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetWindowText(hwnd))

    def get_screenshots(self, file_name_for_screen, path_to_save_screen, statistics_exel):
        print(f'[bold green]In test[/bold green] {file_name_for_screen}')
        Helper.run(path_to_temp_in_test, file_name_for_screen, exel)
        sleep(wait_for_open)
        win32gui.EnumWindows(self.get_coord_exel, self.coordinate)
        page_num = 1
        list_num = 1
        for page in range(int(statistics_exel['num_of_sheets'])):
            num_of_sheet = 1
            num_of_row = statistics_exel[f'{num_of_sheet}_nrows'] / 65
            print(num_of_row)
            pg.hotkey('ctrl', 'home', interval=0.2)
            for pgdwn in range(int(num_of_row)):
                pg.press('pgdn', interval=0.5)
                CompareImage.grab_coordinate_exel(path_to_save_screen,
                                                  file_name_for_screen,
                                                  list_num,
                                                  page_num,
                                                  self.coordinate[0])
                num_of_sheet += 1
                page_num += 1
            pg.hotkey('ctrl', 'pgdn', interval=0.2)
            sleep(wait_for_press)
            list_num += 1
        # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)

    def run_compare_exel(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for file_name in list_of_files:
                file_name_from = file_name.replace(f'.{extension_to}', f'.{extension_from}')
                name_to_for_test = self.preparing_file_names(file_name)
                name_from_for_test = self.preparing_file_names(file_name_from)
                if extension_to == file_name.split('.')[-1]:
                    print(f'[bold green]In test[/bold green] {name_to_for_test}')
                    print(f'[bold green]In test[/bold green] {name_from_for_test}')
                    self.copy(f'{custom_doc_to}{file_name}',
                              f'{path_to_temp_in_test}{name_to_for_test}')
                    self.copy(f'{custom_doc_from}{file_name_from}',
                              f'{path_to_temp_in_test}{name_from_for_test}')

                    statistics_exel = Exel.opener_exel(path_to_temp_in_test, name_from_for_test)
                    print(statistics_exel['num_of_sheets'])

                    if statistics_exel != {}:
                        Helper.delete(f'{tmp_after}*')
                        Helper.delete(f'{tmp_before}*')
                        self.get_screenshots(name_to_for_test, tmp_after,
                                             statistics_exel)

                        self.get_screenshots(name_from_for_test, tmp_before,
                                             statistics_exel)
                        CompareImage(file_name, 100)

            self.delete(path_to_temp_in_test)
