# -*- coding: utf-8 -*-
import math
import os
import traceback
from multiprocessing import Process
from time import sleep

import win32con
import win32gui
import pyautogui as pg
from rich import print

from loguru import logger
from win32com.client import Dispatch

from config import wait_for_press, wait_for_opening
from libs.helpers.compare_image import CompareImage
from libs.helpers.error_handler import CheckErrors


# methods for working with Excel
class Excel:
    def __init__(self, helper):
        self.helper = helper
        self.check_errors = CheckErrors()
        self.coordinate = []
        self.statistics_excel = None
        self.click = self.helper.click
        self.shell = Dispatch("WScript.Shell")

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

    def open_excel_with_cmd(self, tmp_file_name):
        self.helper.run(self.helper.tmp_dir_in_test, tmp_file_name, self.helper.exel)
        sleep(wait_for_opening)

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == "#32770" \
                and self.check_errors.errors[1] == "Microsoft Excel":
            logger.error(f"'an error has occurred while opening the file'. "
                         f"Copied files: {self.helper.converted_file} "
                         f"and {self.helper.source_file} to 'untested'")

            pg.press('enter')
            self.helper.create_dir(self.helper.opener_errors)
            self.helper.copy_to_folder(self.helper.opener_errors)
            self.check_errors.errors.clear()
            return False

        elif not self.check_errors.errors:
            return True

        else:
            logger.debug(f"Error message: {self.check_errors.errors} "
                         f"Filename: {self.helper.converted_file}")
            self.check_errors.errors.clear()
            return False

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == 'NUIDialog' \
                and self.check_errors.errors[1] == "Microsoft Excel":
            pg.press('alt')
            pg.press('right')
            pg.press('enter')
            self.check_errors.errors.clear()

    def events_handler_when_opening_source_file(self):
        win32gui.EnumWindows(self.check_error, self.check_errors.errors)
        if self.check_errors.errors \
                and self.check_errors.errors[0] == "#32770" \
                and self.check_errors.errors[1] == "Microsoft Excel":
            pg.press('alt')
            pg.press('enter')
            self.check_errors.errors.clear()

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

    def get_excel_statistic(self, wb):
        self.statistics_excel = {
            'num_of_sheets': f'{wb.Sheets.Count}',
        }
        try:
            num_of_sheet = 1
            for sh in wb.Sheets:
                ws = wb.Worksheets(sh.Name)
                used = ws.UsedRange
                nrows = used.Row + used.Rows.Count - 1
                ncols = used.Column + used.Columns.Count - 1
                self.statistics_excel[f'{num_of_sheet}_page_name'] = sh.Name
                self.statistics_excel[f'{num_of_sheet}_nrows'] = nrows
                self.statistics_excel[f'{num_of_sheet}_ncols'] = ncols
                num_of_sheet += 1

        except Exception:
            logger.error(f'\nFailed to get full statistics excel from file: {self.helper.converted_file}\n '
                         f'statistics: {self.statistics_excel}')

    def opener_excel(self, file_name):
        error_processing = Process(target=self.check_errors.run_get_error_exel, args=(self.helper.converted_file,))
        error_processing.start()
        try:
            excel = Dispatch("Excel.Application")
            excel.Visible = False
            workbooks = excel.Workbooks.Open(f'{self.helper.tmp_dir_in_test}{file_name}')
            self.get_excel_statistic(workbooks)
            self.close_opener_excel(excel, workbooks)
            print(f"[bold blue]Number of sheets[/]: {self.statistics_excel['num_of_sheets']}")
            return True

        except Exception:
            error = traceback.format_exc()
            logger.error(f'{error} happened while opening file: {self.helper.converted_file}')
            self.statistics_excel = None
            logger.error(f"Can't open file: {self.helper.source_file}. Copied files to 'untested'")
            self.helper.copy_to_folder(self.helper.failed_source)
            return False

        finally:
            error_processing.terminate()

    def close_excel(self):
        pg.hotkey('ctrl', 'z')
        os.system("taskkill /t /im  EXCEL.EXE")
        sleep(0.2)
        self.events_handler_when_closing()

    def close_opener_excel(self, excel, workbooks):
        try:
            workbooks.Close(False)
            excel.Quit()
        except Exception:
            error = traceback.format_exc()
            logger.error(f'{error} happened while closing file: {self.helper.converted_file}')
        finally:
            os.system("taskkill /t /im  EXCEL.EXE")
