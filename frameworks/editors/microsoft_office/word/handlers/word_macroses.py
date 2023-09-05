# -*- coding: utf-8 -*-
from time import sleep
import pyautogui as pg

from frameworks.StaticData import StaticData

from ....key_actions import KeyActions


class WordMacroses:
    @staticmethod
    def prepare_document_for_test():
        KeyActions.click(f"{StaticData.project_dir}/assets/image_templates/word/layout.png")
        sleep(0.3)
        KeyActions.click(f"{StaticData.project_dir}/assets/image_templates/word/transfers.png")
        pg.press('down', interval=0.1)
        pg.press('enter')
        KeyActions.click(f"{StaticData.project_dir}/assets/image_templates/powerpoint/view.png")
        sleep(0.3)
        KeyActions.click(f"{StaticData.project_dir}/assets/image_templates/word/one_page.png")
        KeyActions.click(f"{StaticData.project_dir}/assets/image_templates/word/resolution100.png")
        pg.moveTo(100, 0)
        sleep(0.5)
