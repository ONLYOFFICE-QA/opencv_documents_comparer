# -*- coding: utf-8 -*-
from loguru import logger
from rich import print
import pyautogui as pg

from config import version
from framework.libre_office import LibreOffice
from libs.helpers.fileutils import FileUtils


class OpenerOdp(LibreOffice):
    def run_opener(self, list_of_files):
        logger.info(f'Opener {self.doc_helper.converted_extension} with LibreOffice on version: {version} is running.')
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.endswith((".odp", ".ODP")):
                continue
            self.doc_helper.preparing_files_for_opening_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            self.open_libre_office_with_cmd(self.doc_helper.tmp_converted_file)
            self.errors_handler_when_opening()
            self.close_libre()
            self.doc_helper.tmp_cleaner()
