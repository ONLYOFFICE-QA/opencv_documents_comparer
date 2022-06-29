# -*- coding: utf-8 -*-
import pyautogui as pg
from loguru import logger
from rich import print

from config import version
from framework.libre_office import LibreOffice
from framework.power_point import PowerPoint
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper


class OdpPptxCompare:

    def __init__(self):
        self.helper = Helper('odp', 'pptx')
        self.power_point = PowerPoint(self.helper)
        self.libre = LibreOffice(self.helper)
        logger.info(f'The {self.helper.source_extension} to {self.helper.converted_extension} '
                    f'comparison on version: {version} is running.')

    def run_compare_odp_pptx(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".pptx", ".PPTX")):
                continue
            self.helper.preparing_files_for_test()
            print(f'[bold green]In test[/] {self.helper.converted_file}')
            if not self.power_point.opener_power_point(self.helper.tmp_dir_in_test, self.helper.tmp_name):
                continue
            self.power_point.open_presentation_with_cmd(self.helper.tmp_name_converted_file)
            if not self.power_point.errors_handler_when_opening():
                self.power_point.close_presentation_with_hotkey()
                continue
            self.power_point.get_screenshot(self.helper.tmp_dir_converted_image)
            self.power_point.close_presentation_with_hotkey()

            print(f'[bold green]In test[/] {self.helper.source_file}')
            self.libre.open_libre_office_with_cmd(self.helper.tmp_name_source_file)
            self.libre.get_screenshot_odp(self.helper.tmp_dir_source_image, self.power_point.slide_count)
            self.libre.close_libre()

            CompareImage(self.helper, odp=True)
            self.helper.tmp_cleaner()
