# -*- coding: utf-8 -*-
import math
import subprocess as sb
from os.path import join, isfile
from time import sleep

import pyautogui as pg

import config as config
from frameworks.StaticData import StaticData
from host_control import Window
from frameworks.image_handler import Image

from .handlers import ExcelMacroses, ExcelEvents
from .table_info import TableInfo
from ...editor import Editor


# methods for working with Excel
class Excel(Editor):

    def __init__(self):
        self.path = join(config.ms_office, StaticData.excel)
        self.delay_after_open = config.delay_excel

    @property
    def path(self) -> str:
        return self._path

    @path.setter
    def path(self, value):
        if not isfile(value):
            raise print(f"[bold red]|ERROR|The {type(self).__name__} path does not exist: {value}")
        self._path = value

    @property
    def delay_after_open(self) -> int | float:
        return self._delay_after_open

    @delay_after_open.setter
    def delay_after_open(self, value: int | float):
        if isinstance(value, int) or isinstance(value, float):
            self._delay_after_open = value
        else:
            raise print("[bold red]|ERROR| Delay excel must by integer or float.")

    @property
    def process_names(self) -> list:
        return ['EXCEL.EXE']

    def events_handler(self) -> object:
        return ExcelEvents()

    @property
    def formats(self) -> tuple:
        return '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xls', '.xlt'

    def open(self, file_path: str) -> None:
        sb.Popen(f"{self.path} -t {file_path}")

    def close(self, hwnd: int) -> bool:
        self._undo_changes()
        Window.close(hwnd)
        sleep(0.5)
        return Window().handle_events(ExcelEvents().when_closing)

    def page_amount(self, file_path: str) -> dict:
        return TableInfo(file_path).get()

    def set_size(self, hwnd: int) -> None:
        Window.set_size_max(hwnd)

    def make_screenshots(self, hwnd: int, screen_path: str, page_amount: dict) -> None:
        ExcelMacroses.prepare_for_test()
        coordinate = self.screen_coordinate(Window.get_coordinate(hwnd))
        list_num = 1
        for press in range(int(page_amount['sheets_count'])):
            pg.hotkey('ctrl', 'pgup', interval=0.05)
        for sheet in range(int(page_amount['sheets_count'])):
            pg.hotkey('ctrl', 'home', interval=0.2)
            screen_num = 1
            Image.make_screenshot(join(screen_path, f'list_{list_num}_page_{screen_num}.png'), coordinate)
            if f'{list_num}_nrows' in page_amount:
                num_of_row = page_amount[f'{list_num}_nrows'] / 65
            else:
                num_of_row = 2
            for pgdwn in range(math.ceil(num_of_row)):
                pg.press('pgdn', interval=0.5)
                screen_num += 1
                Image.make_screenshot(join(screen_path, f'list_{list_num}_page_{screen_num}.png'), coordinate)
            pg.hotkey('ctrl', 'pgdn', interval=0.05)
            sleep(0.5)
            list_num += 1

    @staticmethod
    def screen_coordinate(window_coordinate):
        return (
            window_coordinate[0] + 10,
            window_coordinate[1] + 170,
            window_coordinate[2] - 30,
            window_coordinate[3] - 70
        )

    @staticmethod
    def _undo_changes():
        try:
            pg.hotkey('ctrl', 'z', presses=5)
        except pg.FailSafeException as e:
            print(f"[red]|WARNING| Can't undo changes. Exception: {e}")
