# -*- coding: utf-8 -*-
import pyautogui as pg

from framework.StaticData import StaticData


class KeyActions:

    @staticmethod
    def click(image_path):
        try:
            pg.click(f"{StaticData.PROJECT_DIR}/data/image_templates/{image_path}")
            return True
        except TypeError:
            return False
