# -*- coding: utf-8 -*-
import subprocess as sb
from os.path import join, isfile
from time import sleep
import pyautogui as pg
from host_tools import Window

import config
from frameworks.StaticData import StaticData
from frameworks.image_handler import Image

from ..editor import Editor
from .handlers import LibreEvents, LibreMacroses


class LibreOffice(Editor):
    def __init__(self):
        self.path = join(config.libre_office, StaticData.libre)
        self.window = Window()
        self.delay_after_open = config.delay_libre

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
            raise print("[bold red]|ERROR| Delay libre must by integer or float.")

    @property
    def process_names(self) -> list:
        return ['soffice.bin']

    @property
    def formats(self) -> tuple:
        return '.ods', '.odp', '.odt'

    def events_handler(self) -> object:
        return LibreEvents()

    def open(self, file_path: str) -> None:
        sb.Popen(f"{self.path} -o {file_path}")

    def close(self, hwnd: int) -> bool:
        self._undo_changes()
        self.window.close(hwnd)
        sleep(0.5)
        return self.window.handle_events(self.events_handler().when_closing)

    def page_amount(self, file_path: str) -> int:
        ...

    def set_size(self, hwnd: int) -> None:
        self.window.set_size(hwnd, 0, 0, 1920, 1440)

    def make_screenshots(self, hwnd: int, screen_path: str, page_amount: int) -> None:
        LibreMacroses.prepare_for_test()
        coordinate = self.screen_coordinate(self.window.get_coordinate(hwnd))
        for page in range(int(page_amount)):
            Image.make_screenshot(join(screen_path, f'page_{page + 1}.png'), coordinate)
            pg.press('pgdn')
            sleep(0.5)

    @staticmethod
    def screen_coordinate(window_coordinate: list) -> tuple:
        return (
            window_coordinate[0] + 370,
            window_coordinate[1] + 200,
            window_coordinate[2] - 120,
            window_coordinate[3] - 100
        )

    @staticmethod
    def _undo_changes():
        try:
            pg.hotkey('ctrl', 'z', presses=5)
        except pg.FailSafeException as e:
            print(f"[red]|WARNING| Can't undo changes. Exception: {e}")
