# -*- coding: utf-8 -*-
import math
import win32gui
from loguru import logger
from multiprocessing import Process
import configuration as config
import subprocess as sb
import pyautogui as pg
from time import sleep
import win32con
from win32com.client import Dispatch

from data.project_configurator import ProjectConfig
from framework.actions.document_actions import DocActions
from framework.telegram import Telegram
from framework.compare_image import CompareImage
from framework.FileUtils import FileUtils


# methods for working with Excel
class Excel:
    def __init__(self):
        self.doc_helper = ProjectConfig.DOC_ACTIONS
        self.errors = []
        self.statistics_excel = None
        self.windows_handler_number = None
        self.files_with_errors_when_opening = []
        self.num_of_sheets = ''

    def get_windows_title(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if class_name in ['#32770', 'NUIDialog', 'ThunderDFrame']:
                win32gui.SetForegroundWindow(hwnd)
                self.errors.clear()
                self.errors.append(class_name)
                self.errors.append(window_text)

    def error_handler_for_thread(self):
        while True:
            win32gui.EnumWindows(self.get_windows_title, self.errors)
            if self.errors:
                match self.errors:
                    case ['#32770', 'Microsoft Visual Basic']:
                        if not DocActions.click('excel/end.png'):
                            sb.call(f'powershell.exe kill -Name EXCEL', shell=True)
                    case ['#32770', 'Microsoft Excel']:
                        pg.press('enter')
                        sleep(1)
                        pg.press('enter')
                    case ['#32770', 'Monopoly'] | \
                         ['NUIDialog', ('Microsoft Excel' | 'Microsoft Excel - проверка совместимости')]:
                        pg.press('enter')
                    case ['ThunderDFrame', 'Functions List']:
                        pg.hotkey('alt', 'f4')
                    case ['ThunderDFrame', 'Select Players and Times']:
                        pg.press('tab', presses=6, interval=0.2)
                        pg.press('enter', interval=0.2)
                    case _:
                        message = f'New Event: {self.errors} while opening: {self.doc_helper.converted_file}'
                        logger.debug(message)
                        Telegram.send_message(message)
                self.errors.clear()

    def prepare_excel_for_test(self):
        if DocActions.click('/excel/turn_on_content.png'):
            sleep(1)
            win32gui.EnumWindows(self.get_windows_title, self.errors)
            if self.errors:
                self.errors.clear()
                error_processing = Process(target=self.error_handler_for_thread)
                error_processing.start()
                sleep(7)
                error_processing.terminate()

    def open_excel_with_cmd(self, file_name):
        self.errors.clear()
        sb.Popen(f"{config.ms_office}/{ProjectConfig.EXCEL} -t {ProjectConfig.TMP_DIR_IN_TEST}/{file_name}")
        self.waiting_for_opening_excel()

    def check_open_excel(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'XLMAIN' and win32gui.GetWindowText(hwnd) != '':
                self.windows_handler_number = hwnd

    def waiting_for_opening_excel(self):
        self.windows_handler_number = None
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_excel, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(config.wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"Too long to open file: {self.doc_helper.converted_file}")
                self.doc_helper.copy_testing_files_to_folder(self.doc_helper.too_long_to_open_files)
                break

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.get_windows_title, self.errors)
        if self.errors and self.errors[0] == "#32770" and self.errors[1] == "Microsoft Excel":
            logger.error(f"'an error has occurred while opening' when opening file: {self.doc_helper.converted_file}.")
            pg.press('enter', presses=5)
            self.files_with_errors_when_opening.append(self.doc_helper.converted_file)
            self.doc_helper.copy_testing_files_to_folder(self.doc_helper.opener_errors)
            self.errors.clear()
            return False
        elif not self.errors:
            return True
        else:
            logger.debug(f"New Error\nError message: {self.errors}\nFile: {self.doc_helper.converted_file}")
            self.errors.clear()
            return False

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.get_windows_title, self.errors)
        if self.errors and self.errors[0] == 'NUIDialog' and self.errors[1] == "Microsoft Excel":
            pg.press('right')
            pg.press('enter')
            self.errors.clear()

    def events_handler_when_opening_source_file(self):
        win32gui.EnumWindows(self.get_windows_title, self.errors)
        if self.errors and self.errors[0] == "#32770" and self.errors[1] == "Microsoft Excel":
            pg.press('enter')
            self.errors.clear()

    def set_windows_size_excel(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    # gets the coordinates of the window
    # sets the size and position of the window
    def get_coordinate_exel(self):
        coordinate = [win32gui.GetWindowRect(self.windows_handler_number)]
        coordinate = coordinate[0]
        coordinate = (coordinate[0] + 10,
                      coordinate[1] + 170,
                      coordinate[2] - 30,
                      coordinate[3] - 70)
        return coordinate

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshots(self, path_to_save):
        if win32gui.IsWindow(self.windows_handler_number):
            self.set_windows_size_excel()
            coordinate = self.get_coordinate_exel()
            self.prepare_excel_for_test()
            list_num = 1
            for press in range(int(self.statistics_excel['num_of_sheets'])):
                pg.hotkey('ctrl', 'pgup', interval=0.05)
            for sheet in range(int(self.statistics_excel['num_of_sheets'])):
                pg.hotkey('ctrl', 'home', interval=0.2)
                page_num = 1
                CompareImage.grab_coordinate(f"{path_to_save}/list_{list_num}_page_{page_num}.png", coordinate)
                if f'{list_num}_nrows' in self.statistics_excel:
                    num_of_row = self.statistics_excel[f'{list_num}_nrows'] / 65
                else:
                    num_of_row = 2
                for pgdwn in range(math.ceil(num_of_row)):
                    pg.press('pgdn', interval=0.5)
                    page_num += 1
                    CompareImage.grab_coordinate(f"{path_to_save}/list_{list_num}_page_{page_num}.png", coordinate)
                pg.hotkey('ctrl', 'pgdn', interval=0.05)
                sleep(0.5)
                list_num += 1
        else:
            message = f'Invalid window handle when get_screenshot, file: {self.doc_helper.converted_file}'
            Telegram.send_message(message)
            logger.error(message)

    def get_excel_statistic(self, wb):
        self.statistics_excel = {'num_of_sheets': f'{wb.Sheets.Count}'}
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
        except Exception as e:
            massage = f'Failed to get full statistics excel from file: {self.doc_helper.converted_file}\n' \
                      f'statistics: {self.statistics_excel}\nException: {e}'
            logger.error(massage)
            Telegram.send_message(massage)

    def get_information_about_table(self, file_name):
        error_processing = Process(target=self.error_handler_for_thread, args=(self.doc_helper.converted_file,))
        error_processing.start()
        try:
            excel = Dispatch("Excel.Application")
            excel.Visible = False
            workbooks = excel.Workbooks.Open(f'{ProjectConfig.TMP_DIR_IN_TEST}/{file_name}')
            self.get_excel_statistic(workbooks)
            self.close_opener_excel(excel, workbooks)
            self.num_of_sheets = self.statistics_excel['num_of_sheets']
            return True

        except Exception as e:
            logger.error(f'{e} happened while opening file: {self.doc_helper.converted_file} \nException: {e}')
            self.statistics_excel = None
            self.doc_helper.copy_testing_files_to_folder(self.doc_helper.failed_source)
            return False

        finally:
            error_processing.terminate()

    def close_excel(self):
        pg.hotkey('ctrl', 'z')
        FileUtils.run_command(f"taskkill /im  {ProjectConfig.EXCEL}")
        sleep(0.2)
        self.events_handler_when_closing()

    def close_opener_excel(self, excel, workbooks):
        try:
            workbooks.Close(False)
            excel.Quit()
        except Exception as e:
            logger.error(f'{e} happened while closing file: {self.doc_helper.converted_file}\nException: {e}')
        finally:
            FileUtils.run_command("taskkill /t /im  EXCEL.EXE")
