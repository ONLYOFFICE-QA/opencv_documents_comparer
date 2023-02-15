# -*- coding: utf-8 -*-
import subprocess as sb
from multiprocessing import Process
from os.path import join
from time import sleep

import pyautogui as pg
import win32con
import win32gui
from loguru import logger
from rich import print
from win32com.client import Dispatch

import settings as config
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.key_actions import KeyActions
from framework.compare_image import CompareImage
from framework.telegram import Telegram


class PowerPoint:
    def __init__(self):
        self.doc_helper = ProjectConfig.DOC_ACTIONS
        self.errors = []
        self.slide_count = None
        self.windows_handler_number = None
        self.files_with_errors_when_opening = []

    @staticmethod
    def prepare_presentation_for_test():
        KeyActions.click('/excel/turn_on_content.png')
        KeyActions.click('/excel/turn_on_content.png')
        sleep(0.2)
        KeyActions.click('/powerpoint/ok.png')
        KeyActions.click('/powerpoint/view.png')
        pg.click()
        sleep(0.2)
        KeyActions.click('/powerpoint/normal_view.png')
        KeyActions.click('/powerpoint/scale.png')
        pg.press('tab')
        pg.write('100', interval=0.1)
        pg.press('enter')
        sleep(0.5)

    # sets the size and position of the window
    def set_windows_size_pp(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2200, 1420, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    # Checks the window title
    def get_windows_title(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if class_name == '#32770' and window_text == "Microsoft PowerPoint" or class_name == 'NUIDialog':
                win32gui.SetForegroundWindow(hwnd)
                self.errors.clear()
                self.errors.append(class_name)
                self.errors.append(window_text)

    def error_handler_for_thread(self, filename):
        while True:
            win32gui.EnumWindows(self.get_windows_title, self.errors)
            if self.errors:
                match self.errors:
                    case ['NUIDialog', 'Пароль']:
                        logger.error(f'"{self.errors}" happened while opening: {self.doc_helper.converted_file}')
                        pg.press('right', presses=2)
                        pg.press('enter')
                    case _:
                        massage = f'"New Event {self.errors}" happened while opening: {self.doc_helper.converted_file}'
                        logger.debug(massage)
                        Telegram.send_message(massage)
            self.errors.clear()

    def get_slide_count(self):
        error_processing = Process(target=self.error_handler_for_thread)
        error_processing.start()
        presentation = Dispatch("PowerPoint.application")
        try:
            presentation = presentation.Presentations.Open(file_path)
            self.slide_count = len(presentation.Slides)
            print(f"[bold blue]Number of Slides[/]:{self.slide_count}")
            return True

        except Exception as e:
            logger.error(f'Exception while opening presentation. {self.doc_helper.converted_file}\nException: {e}')
            self.slide_count = None
            self.doc_helper.copy_testing_files_to_folder()
            return False

        finally:
            error_processing.terminate()
            self.close_presentation_with_hotkey()
            FileUtils.run_command(f"taskkill /im {ProjectConfig.POWERPOINT}")

    def open_presentation_with_cmd(self, file_path):
        self.errors.clear()
        sb.Popen(f"{config.ms_office}/{ProjectConfig.POWERPOINT} -t {file_path}")
        self.waiting_for_opening_power_point()

    def check_open_power_point(self, hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'PPTFrameClass' and win32gui.GetWindowText(hwnd) != '':
                self.windows_handler_number = hwnd

    def waiting_for_opening_power_point(self):
        self.windows_handler_number = None
        stop_waiting = 1
        while True:
            win32gui.EnumWindows(self.check_open_power_point, self.windows_handler_number)
            if self.windows_handler_number:
                sleep(config.wait_for_opening)
                break
            sleep(0.5)
            stop_waiting += 1
            if stop_waiting == 1000:
                logger.error(f"'Too long to open file: {self.doc_helper.converted_file} ")
                self.doc_helper.copy_testing_files_to_folder()
                break

    def errors_handler_when_opening(self):
        win32gui.EnumWindows(self.get_windows_title, self.errors)
        if self.errors and self.errors[0] == 'NUIDialog':
            pg.press('enter')
            sleep(2)
            self.errors.clear()
        elif self.errors and self.errors[0] == "#32770" and self.errors[1] == "Microsoft PowerPoint":
            logger.error(f"'an error has occurred while opening the file' Files: {self.doc_helper.converted_file}")
            self.files_with_errors_when_opening.append(self.doc_helper.converted_file)
            pg.press('esc', presses=3, interval=0.2)
            self.doc_helper.copy_testing_files_to_folder(self.doc_helper.opener_errors)
            self.errors.clear()
            return False
        elif not self.errors:
            return True
        else:
            logger.debug(f"New Error\nError message: {self.errors}\nFile: {self.doc_helper.converted_file}")
            self.doc_helper.copy_testing_files_to_folder()
            return False

    def events_handler_when_closing(self):
        win32gui.EnumWindows(self.get_windows_title, self.errors)
        if self.errors and self.errors[0] == 'NUIDialog':
            pg.press('right')
            pg.press('enter')
            self.errors.clear()

    # gets the coordinates of the window
    def get_coordinate_pp(self):
        coordinate = [win32gui.GetWindowRect(self.windows_handler_number)]
        coordinate = coordinate[0]
        coordinate = (coordinate[0] + 350,
                      coordinate[1] + 170,
                      coordinate[2] - 120,
                      coordinate[3] - 100)
        return coordinate

    # opens the document
    # takes a screenshot by coordinates
    def get_screenshot(self, path_to_save_screen):
        if win32gui.IsWindow(self.windows_handler_number):
            self.set_windows_size_pp()
            coordinate = self.get_coordinate_pp()
            self.prepare_presentation_for_test()
            page_num = 1
            for page in range(self.slide_count):
                CompareImage.grab_coordinate(join(path_to_save_screen, f"page_{page_num}.png"), coordinate)
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {self.doc_helper.converted_file}'
            Telegram.send_message(massage)
            logger.error(massage)

    def close_presentation_with_hotkey(self):
        if win32gui.IsWindow(self.windows_handler_number):
            win32gui.SetForegroundWindow(self.windows_handler_number)
        pg.hotkey('ctrl', 'z', interval=0.2)
        pg.hotkey('ctrl', 'q', interval=0.2)
        sleep(0.2)
        self.events_handler_when_closing()
