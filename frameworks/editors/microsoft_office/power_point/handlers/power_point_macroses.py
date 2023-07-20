# -*- coding: utf-8 -*-
from os.path import join
from time import sleep

import pyautogui as pg

from frameworks.StaticData import StaticData
from frameworks.host_control import KeyActions


class PowerPointMacroses:
    IMAGE_TEMPLATES_DIR = join(StaticData.project_dir, 'assets', 'image_templates', 'powerpoint')

    @classmethod
    def prepare_for_test(cls):
        KeyActions.click(f"{cls.IMAGE_TEMPLATES_DIR}/turn_on_content.png")
        KeyActions.click(f"{cls.IMAGE_TEMPLATES_DIR}/turn_on_content.png")
        sleep(0.2)
        KeyActions.click(f"{cls.IMAGE_TEMPLATES_DIR}/ok.png")
        KeyActions.click(f"{cls.IMAGE_TEMPLATES_DIR}/view.png")
        sleep(0.2)
        KeyActions.click(f"{cls.IMAGE_TEMPLATES_DIR}/normal_view.png")
        KeyActions.click(f"{cls.IMAGE_TEMPLATES_DIR}/scale.png")
        pg.press('tab')
        pg.write('100', interval=0.1)
        pg.press('enter')
        sleep(0.5)
