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
            if self.helper.converted_file.endswith((".pptx", ".PPTX")):
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test[/] {self.helper.converted_file}')
                self.power_point.opener_power_point(self.helper.tmp_dir_in_test, self.helper.tmp_name)

                if self.power_point.slide_count is not None:
                    self.power_point.open_presentation_with_cmd(self.helper.tmp_name_converted_file)
                    if self.power_point.check_errors.errors \
                            and self.power_point.check_errors.errors[0] == "#32770" \
                            and self.power_point.check_errors.errors[1] == "Microsoft PowerPoint":

                        logger.error(f"'an error has occurred while opening the file'. "
                                     f"Copied files: {self.helper.converted_file} "
                                     f"and {self.helper.source_file} to 'failed_to_open_converted_file'")

                        pg.press('esc', presses=3, interval=0.2)
                        self.power_point.close_presentation_with_hotkey()
                        self.helper.copy_to_folder(self.helper.untested_folder)
                        self.power_point.check_errors.errors.clear()

                    elif not self.power_point.check_errors.errors:
                        self.power_point.get_screenshot(self.helper.tmp_dir_converted_image)
                        self.power_point.close_presentation_with_hotkey()

                        print(f'[bold green]In test[/] {self.helper.source_file}')
                        self.libre.open_libre_office_with_cmd(self.helper.tmp_name_source_file)
                        self.libre.get_screenshot_odp(self.helper.tmp_dir_source_image, self.power_point.slide_count)
                        self.libre.close_libre()

                        CompareImage(self.helper, odp=True)

                    else:
                        logger.debug(f"Error message: {self.power_point.check_errors.errors} "
                                     f"Filename: {self.helper.converted_file}")

                else:
                    logger.error(f"Can't open {self.helper.source_file}. Copied files"
                                 " to 'failed_to_open_file'")

                    self.helper.copy_to_folder(self.helper.failed_source)
            self.helper.tmp_cleaner()
