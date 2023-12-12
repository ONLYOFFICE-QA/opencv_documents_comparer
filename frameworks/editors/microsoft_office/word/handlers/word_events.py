# -*- coding: utf-8 -*-
import pyautogui as pg
from os.path import join, dirname, realpath
from rich import print
from time import sleep

from frameworks.decorators import singleton
from host_tools import Window, File

from ....events import Events


@singleton
class WordEvents(Events):

    def __init__(self):
        self.warning_windows = File.read_json(join(dirname(realpath(__file__)), 'word_warning_window.json'))

    @property
    def window_class_names(self) -> list:
        return ['OpusApp', "#32770", 'bosa_sdm_msword', 'ThunderDFrame', 'NUIDialog', "MsoSplash"]

    def _warning_window(self, hwnd: int) -> bool:
        for _, info in self.warning_windows.items():
            if Window.get_window_info(hwnd, info['window_title'], info['window_text']):
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| {info['message']}\n{'-' * 90}")
                _button_info = Window.get_window_info(hwnd, info['button_title'], info['button_name'])[0]
                Window.click_on_button(_button_info)
                return True
        return False

    def when_opening(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ['#32770', 'Microsoft Word']:
                _hwnd = Window.get_hwnd('#32770', 'Microsoft Word') if not hwnd else hwnd
                if self._warning_window(_hwnd):
                    raise
                print(f"[bold red]\n{'-' * 90}\n|ERROR| an error has occurred while opening the file\n{'-' * 90}")
                Window.close(_hwnd)
                return True

            case ['bosa_sdm_msword', 'Преобразование файла']:
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Word File conversion\n{'-' * 90}")
                Window.close(hwnd)

            case [_, 'Пароль']:
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Enter password\n{'-' * 90}")
                pg.press('tab')
                pg.press('enter')

        return False

    def when_closing(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ["NUIDialog", "Microsoft Word"]:
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW|Save file\n{'-' * 90}")
                pg.press('right')
                pg.press('enter')

            case ["#32770", "Microsoft Word"]:
                print(f"operation aborted")
                pg.press('enter')

    @staticmethod
    def handler(class_name, windows_text, hwnd: int = None):
        match [class_name, windows_text]:

            case ['#32770', 'Microsoft Word'] | \
                 ['NUIDialog', 'Извещение системы безопасности Microsoft Word']:
                Window.close(hwnd)
                pg.press('left')
                pg.press('enter')

            case ['#32770', 'Microsoft Visual Basic for Applications'] | \
                 ['bosa_sdm_msword', 'Преобразование файла'] | \
                 ['NUIDialog', 'Microsoft Word']:
                Window.close(hwnd)
                pg.press('enter')

            case ['bosa_sdm_msword', 'Пароль']:
                pg.press('tab')
                pg.press('enter')

            case ['#32770', 'Удаление нескольких элементов']:
                print(class_name, windows_text)

            case ['#32770', 'Сохранить как']:
                Window.close(hwnd)

            case ['bosa_sdm_msword', 'Показать исправления']:
                sleep(2)
                pg.press('tab', presses=3)
                pg.press('enter')
