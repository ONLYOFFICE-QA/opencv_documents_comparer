# -*- coding: utf-8 -*-

import os

import pyautogui as pg
from loguru import logger
from rich import print

from libs.functional.presentation.libre_presentation import LibrePresentation
from libs.functional.presentation.power_point import PowerPoint
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper


class LibreOdp:

    def __init__(self):
        self.helper = Helper('odp', 'pptx')
        self.power_point = PowerPoint(self.helper)
        self.libre = LibrePresentation(self.helper)

    def run_compare_odp_pptx(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".pptx", ".PPTX")):
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test[/] {self.helper.converted_file}')
                self.power_point.opener_power_point(self.helper.tmp_dir_in_test, self.helper.tmp_name)

                if self.power_point.slide_count is not None:
                    self.power_point.open_presentation_with_cmd(self.helper.tmp_name_converted_file)
                    if self.power_point.check_errors.errors \
                            and self.power_point.check_errors.errors[0] == "#32770" \
                            and self.power_point.check_errors.errors[1] == "Microsoft PptPptxCompareImg":

                        logger.error(f"'an error has occurred while opening the file'. "
                                     f"Copied files: {self.helper.converted_file} "
                                     f"and {self.helper.source_file} to 'untested'")

                        pg.press('enter')
                        os.system("taskkill /t /f /im  POWERPNT.EXE")
                        self.helper.copy_to_folder(self.helper.untested_folder)
                        self.power_point.check_errors.errors.clear()

                    elif not self.power_point.check_errors.errors:
                        self.power_point.get_screenshot(self.helper.tmp_dir_converted_image)
                        self.power_point.close_presentation_with_cmd()

                        print(f'[bold green]In test[/] {self.helper.source_file}')
                        self.libre.open_presentation_with_cmd(self.helper.tmp_name_source_file)
                        self.libre.get_screenshot_odp(self.helper.tmp_dir_source_image, self.power_point.slide_count)
                        self.libre.close_presentation()

                        CompareImage(self.helper, libre=True)

                    else:
                        logger.debug(f"Error message: {self.power_point.check_errors.errors} "
                                     f"Filename: {self.helper.converted_file}")

                else:
                    logger.error(f"Can't open {self.helper.source_file}. Copied files"
                                 " to 'failed_to_open_file'")

                    self.helper.copy_to_folder(self.helper.failed_source)
            self.helper.tmp_cleaner()
