# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath

import pyautogui as pg
from rich import print

from frameworks.decorators import singleton
from host_tools import Window, File
from telegram import Telegram

from ....events import Events


@singleton
class PowerPointEvents(Events):

    def __init__(self):
        self.warning_windows = File.read_json(join(dirname(realpath(__file__)), 'pp_warning_window.json'))

    @property
    def window_class_names(self) -> list:
        return ['PPTFrameClass', "#32770", 'NUIDialog', 'MsoSplash']

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

            case ['#32770', 'Microsoft PowerPoint']:
                _hwnd = Window.get_hwnd('#32770', 'Microsoft PowerPoint') if not hwnd else hwnd
                if self._warning_window(_hwnd):
                    raise
                print(f"[bold red]\n{'-' * 90}\n|ERROR| an error has occurred while opening the file\n{'-' * 90}")
                Window.close(_hwnd)
                return True

        return False

    def when_closing(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ['NUIDialog', _]:
                pg.press('right')
                pg.press('enter')

        return False

    def handler_for_thread(self, file_name=None):
        while True:
            errors = Window.get_title_text(self.window_class_names)
            if errors:
                match errors:

                    case ['NUIDialog', 'Пароль']:
                        print(f'[bold red] Enter password: {file_name}')
                        pg.press('right', presses=2)
                        pg.press('enter')

                    case _:
                        massage = f'"New Event {errors}" happened while opening: {file_name}'
                        print(f"[bold red]{massage}"), Telegram().send_message(massage)
