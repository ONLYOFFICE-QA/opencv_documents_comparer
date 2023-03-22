# -*- coding: utf-8 -*-
from os.path import join
import pyautogui as pg
from time import sleep
import win32gui
from loguru import logger

from frameworks.image_handler import CompareImage
from frameworks.telegram import Telegram
from .macroses import Macroses
from frameworks.libre_office.handlers.window import Window


class Screen:
    def __init__(self, windows_handler_number):
        self.windows_handler_number = windows_handler_number
        self.window = Window(self.windows_handler_number)

    def make(self, path_to_save_screen, slide_count, file_name=None):
        if win32gui.IsWindow(self.windows_handler_number):
            self.window.set_size()
            coordinate = self.window.get_coordinate()
            Macroses.prepare_windows_hot_keys()
            page_num = 1
            for page in range(slide_count):
                CompareImage.grab_coordinate(
                    join(path_to_save_screen, f'page_{page_num}.png'),
                    (
                        coordinate[0] + 370,
                        coordinate[1] + 200,
                        coordinate[2] - 120,
                        coordinate[3] - 100
                    )
                )
                pg.press('pgdn')
                sleep(0.5)
                page_num += 1
        else:
            massage = f'Invalid window handle when get_screenshot, File: {file_name}'
            Telegram.send_message(massage)
            logger.error(massage)
