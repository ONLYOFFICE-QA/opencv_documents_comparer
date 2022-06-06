# -*- coding: utf-8 -*-

import os

import pyautogui as pg

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
            if self.helper.converted_file.endswith((".xlsx", ".XLSX")):
                self.helper.preparing_files_for_test()

                if self.helper.converted_file == '1000+Most+Common+Words+in+English+-+Numbers+' \
                                                 '+Vocabulary+for+ESL+EFL+TEFL+TOEFL+TESL+' \
                                                 'English+Learners.xlsx':
                    self.helper.converted_file = '1000MostCommon_renamed.xlsx'

                print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
                self.excel.opener_excel(self.helper.tmp_name)

                if self.excel.statistics_excel is not None:
                    print(f"[bold blue]Number of sheets[/bold blue]: {self.excel.statistics_excel['num_of_sheets']}")

                    self.excel.open_excel_with_cmd(self.helper.tmp_name_converted_file)
                    if not self.excel.check_errors.errors:
                        self.excel.get_screenshots(self.helper.tmp_dir_converted_image)

                        print(f'[bold green]In test[/bold green] {self.helper.source_file}')
                        self.excel.open_excel_with_cmd(self.helper.tmp_name_source_file)
                        if self.excel.check_errors.errors \
                                and self.excel.check_errors.errors[0] == "#32770" \
                                and self.excel.check_errors.errors[1] == "Microsoft Excel":
                            pg.press('alt')
                            pg.press('enter')
                            self.excel.check_errors.errors.clear()

                        self.excel.get_screenshots(self.helper.tmp_dir_source_image)
                        CompareImage(self.helper, koff=99.5)

                    elif self.excel.check_errors.errors \
                            and self.excel.check_errors.errors[0] == "#32770" \
                            and self.excel.check_errors.errors[1] == "Microsoft Excel":
                        logger.error(f"'an error has occurred while opening the file'. "
                                     f"Copied files: {self.helper.converted_file} "
                                     f"and {self.helper.source_file} to 'untested'")

                        pg.press('enter')
                        os.system("taskkill /t /im  EXCEL.EXE")
                        self.helper.copy_to_folder(self.helper.untested_folder)
                        self.excel.check_errors.errors.clear()

                    else:
                        logger.debug(f"Error message: {self.excel.check_errors.errors} "
                                     f"Filename: {self.helper.converted_file}")
                        self.excel.check_errors.errors.clear()

                else:
                    logger.error(f"Can't open file: {self.helper.source_file}. Copied files to 'untested'")
                    self.helper.copy_to_folder(self.helper.failed_source)

            self.helper.tmp_cleaner()
