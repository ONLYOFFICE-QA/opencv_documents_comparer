# -*- coding: utf-8 -*-
import pyautogui as pg
from time import sleep
import config


class LibreMacroses:
    @staticmethod
    def prepare_for_test():
        pg.press('alt', interval=0.1)
        pg.press('right', presses=2, interval=0.1)
        pg.press('down', interval=0.1)
        pg.press('enter', interval=0.1)
        pg.press('alt', interval=0.1)
        pg.press('right', presses=2, interval=0.1)
        pg.press('up', presses=2, interval=0.1)
        pg.press('enter', interval=0.1)
        pg.press('up', interval=0.1)
        pg.press('enter', interval=0.1)
        pg.hotkey('ctrl', 'a', interval=0.1)
        sleep(0.1)
        pg.write('100', interval=0.1)
        pg.press('enter', interval=0.1)
        sleep(0.5)

    @staticmethod
    def close_recovery_window():
        pg.press('esc', interval=0.2)
        pg.press('left', interval=0.2)
        pg.press('enter', interval=0.2)
        sleep(2)
