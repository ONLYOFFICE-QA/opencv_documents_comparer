# -*- coding: utf-8 -*-
import pyautogui as pg
from rich import print

from frameworks.decorators import singleton
from host_control import Window

from .libre_macroses import LibreMacroses
from ...events import Events


@singleton
class LibreEvents(Events):
    @property
    def window_class_names(self) -> list:
        return ['SALFRAME', 'SALSUBFRAME', '#32770', 'bosa_sdm_msword']

    def when_opening(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case ['SALSUBFRAME', 'Ошибка']:
                print(f"[bold red]\n{'-' * 90}\n|ERROR| an error has occurred while opening the file\n{'-' * 90}")
                return True

            case ['SALFRAME', text] if "Восстановление документа LibreOffice" in text:
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Recovery document\n{'-' * 90}")
                LibreMacroses.close_recovery_window()

            case ['bosa_sdm_msword', 'Преобразование файла']:
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Word File conversion\n{'-' * 90}")
                Window.close(hwnd)

            case [_, "Отчёт о сбое"]:
                print(f"[red]\n{'-' * 90}\n|WARNING WINDOW| Failure report\n{'-' * 90}")
                pg.press('esc', interval=0.5)

            case ['SALSUBFRAME', 'Предупреждение']:
                pg.press('enter', interval=0.5)

        return False

    def when_closing(self, class_name, windows_text, hwnd: int = None) -> bool:
        match [class_name, windows_text]:

            case [_, "Сохранить документ?"]:
                pg.press('right')
                pg.press('enter')

        return False
