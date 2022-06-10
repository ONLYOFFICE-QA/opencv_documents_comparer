# -*- coding: utf-8 -*-
import os

import pyautogui as pg
from loguru import logger
from rich import print

from config import version
from framework.excel import Excel
from libs.helpers.helper import Helper


class OpenerXlsx:
    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'xlsx')
        self.excel = Excel(self.helper)
        logger.info(f'Opener {self.helper.converted_extension} with ms Excel on version: {version} is running.')

    def run_opener_xlsx(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                if self.helper.converted_file in self.helper.exception_files["xlsx_files_opening_error_named_range"]:
                    # named range https://bugzilla.onlyoffice.com/show_bug.cgi?id=52628
                    continue
                if self.helper.converted_file in self.helper.exception_files["xlsx_files_opening_error_shape"]:
                    # https://bugzilla.onlyoffice.com/show_bug.cgi?id=57544
                    continue

                self.helper.preparing_files_for_test()

                if self.helper.converted_file == '1000+Most+Common+Words+in+English+-+Numbers+' \
                                                 '+Vocabulary+for+ESL+EFL+TEFL+TOEFL+TESL+' \
                                                 'English+Learners.xlsx':
                    self.helper.converted_file = '1000MostCommon_renamed.xlsx'

                print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
                self.excel.open_excel_with_cmd(self.helper.tmp_name_converted_file)
                if not self.excel.check_errors.errors:
                    self.excel.close_excel()

                elif self.excel.check_errors.errors \
                        and self.excel.check_errors.errors[0] == "#32770" \
                        and self.excel.check_errors.errors[1] == "Microsoft Excel":
                    logger.error(f"'an error has occurred while opening the file'. "
                                 f"Copied files: {self.helper.converted_file} "
                                 f"and {self.helper.source_file} to 'untested'")

                    pg.press('enter')
                    self.excel.close_excel()
                    self.helper.copy_to_folder(self.helper.untested_folder)
                    self.excel.check_errors.errors.clear()
                else:
                    logger.debug(f"Error message: {self.excel.check_errors.errors} "
                                 f"Filename: {self.helper.converted_file}")
                    self.excel.close_excel()
                    self.excel.check_errors.errors.clear()
            self.helper.tmp_cleaner()
