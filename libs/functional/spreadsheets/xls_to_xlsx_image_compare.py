import csv
import io
import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from rich import print

from libs.functional.spreadsheets.xls_to_xlsx_statistic_compare import Excel
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper
from var import *


class ExcelCompareImage(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        self.delete(f'{tmp_in_test}*')
        self.delete(f'{tmp_converted_image}*')
        self.delete(f'{tmp_source_image}*')
        self.coordinate = []
        self.errors = []
        self.run_compare_exel(list_of_files)

    def get_coord_exel(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'XLMAIN':
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def get_windows_title(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword':
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                win32gui.SetForegroundWindow(hwnd)

                self.errors.clear()
                self.errors.append(win32gui.GetClassName(hwnd))
                self.errors.append(win32gui.GetWindowText(hwnd))

    def get_screenshots(self, file_name_for_screen, path_to_save_screen, statistics_exel):
        print(f'[bold green]In test[/bold green] {file_name_for_screen}')
        Helper.run(tmp_in_test, file_name_for_screen, exel)
        sleep(wait_for_opening)
        win32gui.EnumWindows(self.get_coord_exel, self.coordinate)
        coordinate = self.coordinate[0]
        coordinate = (coordinate[0],
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 20)

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
                                                  coordinate)
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
            for converted_file in list_of_files:
                if converted_file.endswith((".xlsx", ".XLSX")):
                    source_file, tmp_name_converted_file, \
                    tmp_name_source_file, tmp_name = self.preparing_files_for_test(converted_file,
                                                                                   extension_converted,
                                                                                   extension_source)
                    print(f'[bold green]In test[/bold green] {tmp_name_converted_file}')
                    print(f'[bold green]In test[/bold green] {tmp_name_source_file}')
                    statistics_exel = Excel.opener_exel(tmp_in_test, tmp_name)
                    print(statistics_exel['num_of_sheets'])

                    if statistics_exel != {}:
                        Helper.delete(f'{tmp_converted_image}*')
                        Helper.delete(f'{tmp_source_image}*')
                        self.get_screenshots(tmp_name_converted_file, tmp_converted_image,
                                             statistics_exel)

                        self.get_screenshots(tmp_name_source_file, tmp_source_image,
                                             statistics_exel)
                        CompareImage(converted_file, 100)

            self.delete(tmp_in_test)
