# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from framework.libre_office import LibreOffice
from settings import version


class OpenerOds(LibreOffice):
    def run_opener(self, list_of_files):
        logger.info(f'Opener {self.doc_helper.converted_extension} with LibreOffice on version: {version} is running.')
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.lower().endswith(".ods"):
                continue
            if self.doc_helper.converted_file == 'rank_prf-09.ods':
                # https://bugzilla.onlyoffice.com/show_bug.cgi?id=59118
                continue
            self.doc_helper.preparing_files_for_opening_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            self.open_libre_office_with_cmd(self.doc_helper.tmp_converted_file)
            self.errors_handler_when_opening()
            self.close_libre()
            self.doc_helper.tmp_cleaner()
