# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath

import pyautogui as pg
from time import sleep

from frameworks.editors.key_actions import KeyActions


class LibreMacroses:
    @staticmethod
    def prepare_for_test():
        press_interval = 0.1
        pg.press('alt', interval=press_interval)
        pg.press('right', presses=2, interval=press_interval)
        pg.press('down', interval=press_interval)
        pg.press('enter', interval=press_interval)

        pg.press('alt', interval=press_interval)
        pg.press('right', presses=2, interval=press_interval)
        pg.press('up', presses=2, interval=press_interval)
        pg.press('enter', interval=press_interval)

        pg.press('up', interval=press_interval)
        pg.press('enter', interval=press_interval)

        KeyActions.click(join(dirname(realpath(__file__)),  'image_templates', 'custom.png'))
        pg.hotkey('ctrl', 'a', interval=press_interval)
        sleep(0.1)
        pg.write('100', interval=press_interval)
        pg.press('enter', interval=press_interval)
        sleep(0.5)

    @staticmethod
    def close_recovery_window():
        pg.press('esc', interval=0.2)
        pg.press('left', interval=0.2)
        pg.press('enter', interval=0.2)
        sleep(2)
