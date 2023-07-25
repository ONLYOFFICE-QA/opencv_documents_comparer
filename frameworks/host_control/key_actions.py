# -*- coding: utf-8 -*-

class KeyActions:
    @staticmethod
    def click(image_path):
        import pyautogui as pg
        try:
            pg.click(image_path)
            return True
        except TypeError:
            return False
