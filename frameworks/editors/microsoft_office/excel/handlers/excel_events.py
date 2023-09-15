# -*- coding: utf-8 -*-
import pyautogui as pg
from multiprocessing import Process
from time import sleep
from rich import print

from frameworks.decorators import singleton
from host_control import Window
from frameworks.telegram import Telegram
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

    @property
    def window_class_names(self) -> list:
        return ['XLMAIN', "#32770", 'NUIDialog', 'ThunderDFrame', 'MsoSplash']

    def when_opening(self, class_name: str, windows_text: str, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ["#32770", "Microsoft Excel"]:
                hwnd = Window.get_hwnd('#32770', 'Microsoft Excel') if not hwnd else hwnd
                if Window.get_window_info(hwnd, 'Static', 'Некоторые формулы содержат циклические ссылки'):
                    print(
                        f"[red]\n{'-' * 90}\n|WARNING WINDOW| Some formulas contain cyclic references. Clicked: 'ОК'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', 'ОК')[0])
                    raise
                elif Window.get_window_info(hwnd, 'Static', 'только для чтения, если вносить изменения не требуется.'):
                    print(
                        f"[red]\n{'-' * 90}\n|WARNING WINDOW| The author advises opening  file for reading only. Clicked: 'ОК'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', 'Д&а')[0])
                    raise
                elif Window.get_window_info(hwnd, 'Static', 'Недопустимая ссылка. Ссылка для названий, значений'):
                    print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Invalid reference. Clicked: 'ОК'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', 'ОК')[0])
                    raise
                print(f"[bold red]\n{'-' * 90}\n|ERROR| an error has occurred while opening the file\n{'-' * 90}")
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
