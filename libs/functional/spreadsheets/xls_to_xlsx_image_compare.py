# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from config import *
from framework.excel import Excel

from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper


class ExcelCompareImage:
    def __init__(self):
        self.helper = Helper('xls', 'xlsx')
        self.excel = Excel(self.helper)
        logger.info(f'The {self.helper.source_extension} to {self.helper.converted_extension} '
                    f'comparison on version: {version} is running.')

    def run_compare_excel_img(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                continue

            self.helper.preparing_files_for_test()

            if self.helper.converted_file == '1000+Most+Common+Words+in+English+-+Numbers+' \
                                             '+Vocabulary+for+ESL+EFL+TEFL+TOEFL+TESL+' \
                                             'English+Learners.xlsx':
                self.helper.converted_file = '1000MostCommon_renamed.xlsx'

            print(f'[bold green]In test[/] {self.helper.converted_file}')
            if not self.excel.opener_excel(self.helper.tmp_name):
                continue
            self.excel.open_excel_with_cmd(self.helper.tmp_name_converted_file)
            if not self.excel.errors_handler_when_opening():
                self.excel.close_excel()
                continue
            self.excel.get_screenshots(self.helper.tmp_dir_converted_image)
            self.excel.close_excel()

            print(f'[bold green]In test[/] {self.helper.source_file}')
            self.excel.open_excel_with_cmd(self.helper.tmp_name_source_file)
            self.excel.events_handler_when_opening_source_file()
            self.excel.get_screenshots(self.helper.tmp_dir_source_image)
            self.excel.close_excel()
            CompareImage(self.helper, koff=99.5)
            self.helper.tmp_cleaner()
