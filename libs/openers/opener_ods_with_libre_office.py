# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from config import version
from framework.libre_office import LibreOffice
from libs.helpers.helper import Helper


class OpenerOds:
    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'ods')
        self.libre = LibreOffice(self.helper)
        self.helper.create_dir(self.helper.opener_errors)
        logger.info(f'Opener {self.helper.converted_extension} with LibreOffice on version: {version} is running.')

    def run_opener_ods(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".ods", ".ODS")):
                continue
            self.helper.preparing_files_for_test()

            print(f'[bold green]In test[/] {self.helper.converted_file}')
            self.libre.open_libre_office_with_cmd(self.helper.tmp_name_converted_file)
            self.libre.errors_handler_when_opening()
            self.libre.close_libre()
            self.helper.tmp_cleaner()
