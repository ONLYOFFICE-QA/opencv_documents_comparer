# -*- coding: utf-8 -*-
from config import version
from data.StaticData import StaticData
from framework.excel import Excel
from libs.helpers.compare_image import CompareImage
from management import *


class ExcelCompareImage(Excel):
    def run_compare(self, list_of_files):
        logger.info(f'The {self.doc_helper.source_extension} to {self.doc_helper.converted_extension} '
                    f'comparison on version: {version} is running.')
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.endswith((".xlsx", ".XLSX")):
                continue
            self.doc_helper.preparing_files_for_compare_test()
            if self.doc_helper.converted_file == '1000+Most+Common+Words+in+English+-+Numbers+' \
                                                 '+Vocabulary+for+ESL+EFL+TEFL+TOEFL+TESL+' \
                                                 'English+Learners.xlsx':
                self.doc_helper.converted_file = '1000MostCommon_renamed.xlsx'

            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            if not self.get_information_about_table(self.doc_helper.tmp_file_for_get_statistic):
                continue
            self.open_excel_with_cmd(self.doc_helper.tmp_converted_file)
            if not self.errors_handler_when_opening():
                self.close_excel()
                continue
            self.get_screenshots(StaticData.TMP_DIR_CONVERTED_IMG)
            self.close_excel()

            print(f'[bold green]In test[/] {self.doc_helper.source_file}')
            self.open_excel_with_cmd(self.doc_helper.tmp_source_file)
            self.events_handler_when_opening_source_file()
            self.get_screenshots(StaticData.TMP_DIR_SOURCE_IMG)
            self.close_excel()

            CompareImage(self.doc_helper, coefficient=100)
            self.doc_helper.tmp_cleaner()
