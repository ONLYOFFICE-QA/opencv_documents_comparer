# -*- coding: utf-8 -*-
import pyautogui as pg
from rich import print

from frameworks.decorators import singleton
from host_tools import Window
from frameworks.telegram import Telegram

from ....events import Events


@singleton
class PowerPointEvents(Events):

    def __init__(self):
        self.window_class = ['#32770', 'NUIDialog']
        self.slides_warning = 'Приложение PowerPoint обнаружило большое количество неиспользуемых образцов слайдов.'

    @property
    def window_class_names(self) -> list:
        return ['PPTFrameClass', "#32770", 'NUIDialog', 'MsoSplash']

    def when_opening(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ['#32770', 'Microsoft PowerPoint']:
                hwnd = Window.get_hwnd('#32770', 'Microsoft PowerPoint') if not hwnd else hwnd
                if Window.get_window_info(hwnd, 'Static', self.slides_warning):
                    print(f"[bold]\n{'-' * 90}\n|WARNING WINDOW| Clear unused slides. Clicked: 'Cancel'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', 'Отмена')[0])
                    raise
                elif Window.get_window_info(hwnd, 'Static', 'Выполнить запуск в безопасном режиме?'):
                    print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Start PP in safe mode. Clicked: 'No'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', '&Нет')[0])
                    raise
                print(f"[bold red]\n{'-' * 90}\n|ERROR| an error has occurred while opening the file\n{'-' * 90}")
                Window.close(hwnd)
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
