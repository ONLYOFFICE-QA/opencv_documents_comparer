# -*- coding: utf-8 -*-
import pyautogui as pg
from rich import print
from os.path import join, dirname, realpath

from frameworks.decorators import singleton
from host_tools import Window, File

from .libre_macroses import LibreMacroses
from ...events import Events


@singleton
class LibreEvents(Events):

    def __init__(self):
        self.warning_windows = File.read_json(join(dirname(realpath(__file__)), 'libre_windows.json'))

    @property
    def window_class_names(self) -> list:
        return ['SALFRAME', 'SALSUBFRAME', '#32770', 'bosa_sdm_msword']

    @staticmethod
    def _warning_handling(warning_type: str, hwnd: int) -> None:
        match warning_type:
            case "Word File conversion":
                Window.close(hwnd) if hwnd else print(f"[red] can't close the window HWND: {hwnd}")
            case "Recovery document":
                LibreMacroses.close_recovery_window()
            case "Failure report":
                pg.press('esc', interval=0.5)
            case "Warning":
                pg.press('enter', interval=0.5)

    def when_opening(self, class_name: str, windows_text: str, hwnd: int = None) -> bool:
        for title, info in self.warning_windows.items():
            info_class_name = info['class_name'] if info['class_name'] else class_name
            if class_name == info_class_name and info['window_text'] in windows_text:
                if info['status'] == "ERROR":
                    print(f"[bold red]\n{'-' * 90}\n|ERROR| {info['message']}\n{'-' * 90}")
                    return True
                elif info['status'] == "WARNING":
                    print(f"[magenta]\n{'-' * 90}\n|WARNING| {info['message']}\n{'-' * 90}")
                    self._warning_handling(title, hwnd)
        return False

    def when_closing(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case [_, "Сохранить документ?"]:
                pg.press('right')
                pg.press('enter')

        return False
