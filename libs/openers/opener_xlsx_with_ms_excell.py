# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from config import version
from framework.excel import Excel
from libs.helpers.helper import Helper


class OpenerXlsx:
    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'xlsx')
        self.excel = Excel(self.helper)
        self.helper.create_dir(self.helper.opener_errors)
        logger.info(f'Opener {self.helper.converted_extension} with ms Excel on version: {version} is running.')

    def run_opener_xlsx(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                continue
            if self.helper.converted_file in self.helper.exception_files["xlsx_files_opening_error_named_range"]:
                # named range https://bugzilla.onlyoffice.com/show_bug.cgi?id=52628
                continue
            self.helper.preparing_files_for_test()
            print(f'[bold green]In test[/] {self.helper.converted_file}')
            self.excel.open_excel_with_cmd(self.helper.tmp_name_converted_file)
            self.excel.errors_handler_when_opening()
            self.excel.close_excel()
            self.helper.tmp_cleaner()
