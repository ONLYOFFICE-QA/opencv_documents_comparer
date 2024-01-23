# -*- coding: utf-8 -*-
import pyautogui as pg
from os.path import join, dirname, realpath
from multiprocessing import Process
from time import sleep
from rich import print

from frameworks.decorators import singleton
from host_tools import Window, File
from telegram import Telegram
from frameworks.StaticData import StaticData
import subprocess as sb

from ....events import Events
from ....key_actions import KeyActions


def handler():
    ExcelEvents().handler_for_thread()


@singleton
class ExcelEvents(Events):

    def __init__(self):
        self.project_dir = StaticData.project_dir
        self.warning_windows = File.read_json(join(dirname(realpath(__file__)), 'excel_warning_window.json'))

    @property
    def window_class_names(self) -> list:
        return ['XLMAIN', "#32770", 'NUIDialog', 'ThunderDFrame', 'MsoSplash']

    def _warning_window(self, hwnd: int) -> bool:
        for _, info in self.warning_windows.items():
            if Window.get_window_info(hwnd, info['window_title'], info['window_text']):
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| {info['message']}\n{'-' * 90}")
                _button_info = Window.get_window_info(hwnd, info['button_title'], info['button_name'])[0]
                Window.click_on_button(_button_info)
                return True
        return False

    def when_opening(self, class_name: str, windows_text: str, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ["#32770", "Microsoft Excel"]:
                _hwnd = Window.get_hwnd('#32770', 'Microsoft Excel') if not hwnd else hwnd
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

    def when_opening_source_file(self):
        events = Window.get_title_text(self.window_class_names)
        if events and events[0] == "#32770" and events[1] == "Microsoft Excel":
            pg.press('enter')

    def handler_for_thread(self, file_name: str = None):
        while True:
            errors = Window.get_title_text(self.window_class_names)
            if errors:
                match errors:

                    case ['#32770', 'Microsoft Visual Basic']:
                        if not KeyActions.click(f"{self.project_dir}/assets/image_templates/editor/end.png"):
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
                        message = f'New Event: {errors} while opening: {file_name}'
                        print(f"[red]{message}"), Telegram().send_message(message)

    def when_turn_on_content(self):
        if Window.get_title_text(self.window_class_names):
            error_processing = Process(target=handler)
            error_processing.start()
            sleep(7)
            error_processing.terminate()
