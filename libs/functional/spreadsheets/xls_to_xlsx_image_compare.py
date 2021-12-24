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
                    and win32gui.GetWindowText(hwnd) == "Microsoft Excel" \
                    or win32gui.GetClassName(hwnd) == 'NUIDialog' \
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
                error_processing = Process(target=self.check_errors.run_get_error_exel,
                                           args=(self.helper.converted_file,))
                error_processing.start()
                sleep(7)
                error_processing.terminate()
                sleep(0.5)
                win32gui.EnumWindows(self.get_coord_exel, self.coordinate)

        except Exception:
            pass

    def close_excel(self):
        pg.hotkey('ctrl', 'z')
        os.system("taskkill /t /im  EXCEL.EXE")
        sleep(0.2)
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == 'NUIDialog' \
                and self.check_errors.errors[1] == "Microsoft Excel":
            pg.press('alt')
            pg.press('right')
            pg.press('enter')
            self.check_errors.errors.clear()
        pass

    def open_excel_with_cmd(self, tmp_file_name):
        self.helper.run(self.helper.tmp_dir_in_test, tmp_file_name, self.helper.exel)
        sleep(wait_for_opening)
        # check errors
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, path_to_save_screen):
        self.prepare_excel_windows()
        list_num = 1
        win32gui.EnumWindows(self.get_coord_exel, self.coordinate)
        coordinate = self.coordinate[0]
        coordinate = (coordinate[0] + 10,
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 70)

        for press in range(int(self.statistics_excel['num_of_sheets'])):
            pg.hotkey('ctrl', 'pgup', interval=0.05)
        for sheet in range(int(self.statistics_excel['num_of_sheets'])):
            pg.hotkey('ctrl', 'home', interval=0.2)
            page_num = 1
            CompareImage.grab_coordinate_exel(path_to_save_screen,
                                              list_num,
                                              page_num,
                                              coordinate)

            if f'{list_num}_nrows' in self.statistics_excel:
                num_of_row = self.statistics_excel[f'{list_num}_nrows'] / 65
            else:
                num_of_row = 2

            for pgdwn in range(math.ceil(num_of_row)):
                pg.press('pgdn', interval=0.5)
                page_num += 1
                CompareImage.grab_coordinate_exel(path_to_save_screen,
                                                  list_num,
                                                  page_num,
                                                  coordinate)

            pg.hotkey('ctrl', 'pgdn', interval=0.05)
            list_num += 1
            sleep(wait_for_press)
        self.close_excel()

    def run_compare_excel_img(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                self.helper.preparing_files_for_test()

                if self.helper.converted_file == '1000+Most+Common+Words+in+English+-+Numbers+' \
                                                 '+Vocabulary+for+ESL+EFL+TEFL+TOEFL+TESL+' \
                                                 'English+Learners.xlsx':
                    self.helper.converted_file = '1000MostCommon_renamed.xlsx'

                print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
                self.opener_excel(self.helper.tmp_name)

                if self.statistics_excel is not None:
                    print(f"[bold blue]Number of sheets[/bold blue]: {self.statistics_excel['num_of_sheets']}")

                    self.open_excel_with_cmd(self.helper.tmp_name_converted_file)
                    if not self.check_errors.errors:
                        self.get_screenshots(self.helper.tmp_dir_converted_image)

                        print(f'[bold green]In test[/bold green] {self.helper.source_file}')
                        self.open_excel_with_cmd(self.helper.tmp_name_source_file)
                        if self.check_errors.errors \
                                and self.check_errors.errors[0] == "#32770" \
                                and self.check_errors.errors[1] == "Microsoft Excel":
                            pg.press('alt')
                            pg.press('enter')
                            self.check_errors.errors.clear()

                        self.get_screenshots(self.helper.tmp_dir_source_image)
                        CompareImage(self.helper, koff=99.5)

                    elif self.check_errors.errors \
                            and self.check_errors.errors[0] == "#32770" \
                            and self.check_errors.errors[1] == "Microsoft Excel":
                        logger.error(f"'an error has occurred while opening the file'. "
                                     f"Copied files: {self.helper.converted_file} "
                                     f"and {self.helper.source_file} to 'untested'")

                        pg.press('enter')
                        os.system("taskkill /t /im  EXCEL.EXE")
                        self.helper.copy_to_folder(self.helper.untested_folder)
                        self.check_errors.errors.clear()

                    else:
                        logger.debug(f"Error message: {self.check_errors.errors} "
                                     f"Filename: {self.helper.converted_file}")
                        self.check_errors.errors.clear()

                else:
                    logger.error(f"Can't open file: {self.helper.source_file}. Copied files to 'untested'")
                    self.helper.copy_to_folder(self.helper.failed_source)

            self.helper.tmp_cleaner()
