# -*- coding: utf-8 -*-
import pyautogui as pg
from loguru import logger
from rich import print

from config import version
from framework.power_point import PowerPoint
from libs.helpers.helper import Helper


class OpenerPptx:

    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'pptx')
        self.power_point = PowerPoint(self.helper)
        self.helper.create_dir(self.helper.opener_errors)
        logger.info(f'Opener {self.helper.converted_extension} with ms PowerPoint on version: {version} is running.')

    def run_opener(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".pptx", ".PPTX")):
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test[/] {self.helper.converted_file}')
                self.power_point.open_presentation_with_cmd(self.helper.tmp_name_converted_file)
                if self.power_point.check_errors.errors \
                        and self.power_point.check_errors.errors[0] == "#32770" \
                        and self.power_point.check_errors.errors[1] == "Microsoft PowerPoint":

                    logger.error(f"'an error has occurred while opening the file'. "
                                 f"Copied files: {self.helper.converted_file} "
                                 f"and {self.helper.source_file} to 'failed_to_open_converted_file'")

                    pg.press('esc', presses=3, interval=0.2)
                    self.power_point.close_presentation_with_hotkey()
                    self.helper.copy_to_folder(self.helper.opener_errors)
                    self.power_point.check_errors.errors.clear()

                elif not self.power_point.check_errors.errors:
                    self.power_point.close_presentation_with_hotkey()

                else:
                    logger.debug(f"New Error "
                                 f"Error message: {self.power_point.check_errors.errors} "
                                 f"Filename: {self.helper.converted_file}")
                    self.helper.copy_to_folder(self.helper.failed_source)
            self.helper.tmp_cleaner()
