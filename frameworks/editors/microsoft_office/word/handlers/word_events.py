# -*- coding: utf-8 -*-
import pyautogui as pg
from rich import print
from time import sleep

from frameworks.decorators import singleton
from frameworks.host_control import Window

from ....events import Events


@singleton
class WordEvents(Events):

    @property
    def window_class_names(self) -> list:
        return ['OpusApp', "#32770", 'bosa_sdm_msword', 'ThunderDFrame', 'NUIDialog', "MsoSplash"]

    def when_opening(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ['#32770', 'Microsoft Word']:
                hwnd = Window.get_hwnd('#32770', 'Microsoft Word') if not hwnd else hwnd
                if Window.get_window_info(hwnd, 'Static', 'Выполнить запуск в безопасном режиме?'):
                    print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Start Word in safe mode. Clicked: 'No'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', '&Нет')[0])
                    raise
                elif Window.get_window_info(hwnd, 'Static', 'Средство расстановки переносов для приложения Word'):
                    print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Wraparound tool for Word application\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', 'ОК')[0])
                    raise
                elif Window.get_window_info(hwnd, 'Static',
                                            'Документ содержит связи с другими файлами. Обновить в документе данные'):
                    print(
                        f"[red]\n{'-' * 90}\n|WARNING WINDOW| The document contains links to other files. Clicked: 'No'\n{'-' * 90}")
                    Window.click_on_button(Window.get_window_info(hwnd, 'Button', 'Н&ет')[0])
                    raise
                print(f"[bold red]\n{'-' * 90}\n|ERROR| an error has occurred while opening the file\n{'-' * 90}")
                Window.close(hwnd)
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
