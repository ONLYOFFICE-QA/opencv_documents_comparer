# -*- coding: utf-8 -*-
from loguru import logger
from rich import print
import pyautogui as pg

from config import version
from framework.libre_office import LibreOffice
from libs.helpers.helper import Helper


class OpenerOdp:
    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'odp')
        self.libre = LibreOffice(self.helper)
        self.helper.create_dir(self.helper.opener_errors)
        logger.info(f'Opener {self.helper.converted_extension} with LibreOffice on version: {version} is running.')

    def run_opener_odp(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".odp", ".ODP")):
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test[/] {self.helper.converted_file}')
                self.libre.open_libre_office_with_cmd(self.helper.tmp_name_converted_file)
                if self.libre.check_errors.errors:
                    logger.error(f"'an error has occurred while opening the file'. "
                                 f"Copied files: {self.helper.converted_file} "
                                 f"and {self.helper.source_file} to 'failed_to_open_converted_file'")

                    self.helper.copy_to_folder(self.helper.opener_errors)
                    pg.press('enter')
                    self.libre.check_errors.errors.clear()
                    self.libre.close_libre()

                elif not self.libre.check_errors.errors:
                    self.libre.close_libre()

                else:
                    logger.debug(f"Error message: {self.libre.check_errors.errors} "
                                 f"Filename: {self.helper.converted_file}")
            self.helper.tmp_cleaner()
