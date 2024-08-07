# -*- coding: utf-8 -*-
import subprocess as sb
import pyautogui as pg
from os.path import join, isfile
from time import sleep
from rich import print

import config as config
from frameworks.StaticData import StaticData
from host_tools import Window
from frameworks.image_handler import Image

from .document_info import DocumentInfo
from ...editor import Editor
from .handlers import WordEvents, WordMacroses


# methods for working with Word
class Word(Editor):

    def __init__(self):
        self.path = join(config.ms_office, StaticData.word)
        self.delay_after_open = config.delay_word

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
            raise print("[bold red]|ERROR| Delay word must by integer or float.")

    @property
    def process_names(self) -> list:
        return ['WINWORD.EXE']

    @property
    def formats(self) -> tuple:
        return '.doc', '.docm', '.docx', '.dot', '.dotm', '.dotx', '.rtf', '.docxf'

    def events_handler(self) -> object:
        return WordEvents()

    def open(self, file_path) -> None:
        sb.Popen(f"{self.path} -t {file_path}")

    def close(self, hwnd: int) -> bool:
        self._undo_changes()
        Window.close(hwnd)
        sleep(0.5)
        return Window().handle_events(WordEvents().when_closing)

    def page_amount(self, file_path) -> int:
        return DocumentInfo(file_path).page_amount()

    def set_size(self, hwnd: int) -> None:
        Window.set_size(hwnd, 0, 0, 1920, 1440)

    def make_screenshots(self, hwnd: int, screen_path: str, page_amount: int) -> None:
        WordMacroses.prepare_document_for_test()
        coordinate = self.screen_coordinate(Window.get_coordinate(hwnd))
        for page in range(int(page_amount) if page_amount else 1):
            Image.make_screenshot(join(screen_path, f"page_{page + 1}.png"), coordinate)
            pg.press('pgdn')
            sleep(0.5)

    @staticmethod
    def screen_coordinate(window_coordinate):
        return (
            window_coordinate[0],
            window_coordinate[1] + 170,
            window_coordinate[2] - 30,
            window_coordinate[3] - 20
        )

    @staticmethod
    def _undo_changes():
        try:
            pg.hotkey('ctrl', 'z', presses=5)
        except pg.FailSafeException as e:
            print(f"[red]|WARNING| Can't undo changes. Exception: {e}")
