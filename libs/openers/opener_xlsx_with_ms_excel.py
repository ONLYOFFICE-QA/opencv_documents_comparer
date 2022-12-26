# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from configuration import version
from data.project_configurator import ProjectConfig
from framework.excel import Excel


class OpenerXlsx(Excel):
    def run_opener(self, list_of_files):
        logger.info(f'Opener {self.doc_helper.converted_extension} with ms Excel on version: {version} is running.')
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.lower().endswith(".xlsx"):
                continue
            if self.doc_helper.converted_file in ProjectConfig.EXCEPTION_FILES["xlsx_files_opening_error_named_range"]:
                # named range https://bugzilla.onlyoffice.com/show_bug.cgi?id=52628
                continue
            self.doc_helper.preparing_files_for_opening_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            self.open_excel_with_cmd(self.doc_helper.tmp_converted_file)
            self.errors_handler_when_opening()
            self.close_excel()
            self.doc_helper.tmp_cleaner()
