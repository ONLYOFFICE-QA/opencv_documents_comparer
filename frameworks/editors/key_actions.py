# -*- coding: utf-8 -*-
from rich import print

class KeyActions:
    @staticmethod
    def click(image_path):
        import pyautogui as pg
        try:
            pg.click(image_path)
            return True
        except (pg.ImageNotFoundException, TypeError):
            print(f"[red]|ERROR| Can't click on image: {image_path}")
            return False
