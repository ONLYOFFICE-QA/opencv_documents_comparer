# -*- coding: utf-8 -*-
import math
import os
from multiprocessing import Process
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from rich import print

from config import *
from libs.functional.spreadsheets.xls_to_xlsx_statistic_compare import Excel
from libs.helpers.compare_image import CompareImage

source_extension = 'xls'
converted_extension = 'xlsx'


class ExcelCompareImage(Excel):

    # gets the coordinates of the window
    # sets the size and position of the window
    def get_coord_exel(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'XLMAIN':
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.coordinate.clear()
                self.coordinate.append(win32gui.GetWindowRect(hwnd))

    def check_error(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770' \
                    and win32gui.GetWindowText(hwnd) == "Microsoft Excel":
                self.shell.SendKeys('%')
                win32gui.SetForegroundWindow(hwnd)
                sleep(0.5)
                self.check_errors.errors.clear()
                self.check_errors.errors.append(win32gui.GetClassName(hwnd))
                self.check_errors.errors.append(win32gui.GetWindowText(hwnd))

    def prepare_excel_windows(self):
        try:
            pg.click('libs/image_templates/excel/turn_on_content.png')
            sleep(1)

            win32gui.EnumWindows(self.check_errors.get_windows_title, self.check_errors.errors)
            if self.check_errors.errors:
                self.check_errors.errors.clear()
                error_processing = Process(target=self.check_errors.run_get_error_exel, args=(self.file_name_for_log,))
                error_processing.start()
                sleep(7)
                error_processing.terminate()
                sleep(0.5)
                win32gui.EnumWindows(self.get_coord_exel, self.coordinate)

        except Exception:
            print("turn_on_content.png not found")
            pass

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, tmp_file_name, path_to_save_screen, statistics_exel):
        self.helper.run(self.helper.tmp_dir_in_test, tmp_file_name, self.helper.exel)
        sleep(wait_for_opening)
        win32gui.EnumWindows(self.get_coord_exel, self.coordinate)
        coordinate = self.coordinate[0]
        coordinate = (coordinate[0] + 10,
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 70)
        # check errors
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if not self.check_errors.errors:
            self.prepare_excel_windows()
            page_num = 1
            list_num = 1
            for press in range(int(statistics_exel['num_of_sheets'])):
                pg.hotkey('ctrl', 'pgup', interval=0.05)
            for page in range(int(statistics_exel['num_of_sheets'])):
                num_of_sheet = 1
                pg.hotkey('ctrl', 'home', interval=0.2)
                CompareImage.grab_coordinate_exel(path_to_save_screen,
                                                  list_num,
                                                  page_num,
                                                  coordinate)
                page_num += 1
                if f'{num_of_sheet}_nrows' in statistics_exel:
                    num_of_row = statistics_exel[f'{num_of_sheet}_nrows'] / 65
                else:
                    logger.error(f'On {num_of_sheet} sheet, '
                                 f'the number of lines is not found in file {self.file_name_for_log}')
                    num_of_row = 2

                for pgdwn in range(math.ceil(num_of_row)):
                    pg.press('pgdn', interval=0.5)
                    CompareImage.grab_coordinate_exel(path_to_save_screen,
                                                      list_num,
                                                      page_num,
                                                      coordinate)
                    num_of_sheet += 1
                    page_num += 1
                pg.hotkey('ctrl', 'pgdn', interval=0.05)
                sleep(wait_for_press)
                list_num += 1
            os.system("taskkill /t /im  EXCEL.EXE")

    def run_compare_excel_img(self, list_of_files):
        for converted_file in list_of_files:
            if converted_file.endswith((".xlsx", ".XLSX")):
                source_file, tmp_name_converted_file, \
                tmp_name_source_file, tmp_name = self.helper.preparing_files_for_test(converted_file,
                                                                                      converted_extension,
                                                                                      source_extension)
                if converted_file == '1000+Most+Common+Words+in+English+-+Numbers+' \
                                     '+Vocabulary+for+ESL+EFL+TEFL+TOEFL+TESL+' \
                                     'English+Learners.xlsx':
                    converted_file = '1000MostCommon_renamed.xlsx'

                self.file_name_for_log = converted_file

                print(f'[bold green]In test[/bold green] {converted_file}')
                statistics_exel = self.opener_excel(self.helper.tmp_dir_in_test, tmp_name)

                if statistics_exel != {}:
                    print(f"[bold blue]Number of sheets[/bold blue]: {statistics_exel['num_of_sheets']}")
                    self.get_screenshots(tmp_name_converted_file,
                                         self.helper.tmp_dir_converted_image,
                                         statistics_exel)

                    if not self.check_errors.errors:
                        print(f'[bold green]In test[/bold green] {source_file}')
                        self.get_screenshots(tmp_name_source_file,
                                             self.helper.tmp_dir_source_image,
                                             statistics_exel)

                        CompareImage(converted_file, self.helper, koff=100)

                    elif self.check_errors.errors \
                            and self.check_errors.errors[0] == "#32770" \
                            and self.check_errors.errors[1] == "Microsoft Excel":
                        logger.error(f"'an error has occurred while opening the file'. "
                                     f"Copied files: {converted_file} and {source_file} to 'untested'")

                        pg.press('enter')
                        os.system("taskkill /t /im  EXCEL.EXE")
                        self.helper.copy_to_folder(converted_file,
                                                   source_file,
                                                   self.helper.untested_folder)
                        self.check_errors.errors.clear()
                    else:
                        logger.debug(f"Error message: {self.check_errors.errors} Filename: {converted_file}")
                        self.check_errors.errors.clear()

                else:
                    logger.error(f"Can't open file: {source_file}. Copied files to 'untested'")
                    self.helper.copy_to_folder(converted_file,
                                               source_file,
                                               self.helper.untested_folder)

            self.helper.tmp_cleaner()
