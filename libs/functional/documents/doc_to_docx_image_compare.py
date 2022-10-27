# -*- coding: utf-8 -*-
import config
from data.StaticData import StaticData
from framework.word import Word
from framework.compare_image import CompareImage
from management import *


class DocDocxCompareImg(Word):
    def run_compare(self, list_of_files):
        logger.info(f'The {self.doc_helper.source_extension} to {self.doc_helper.converted_extension} '
                    f'comparison on version: {config.version} is running.')

        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.endswith((".docx", ".DOCX")):
                continue
            self.doc_helper.preparing_files_for_compare_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            if not self.get_information_about_document(self.doc_helper.tmp_file_for_get_statistic):
                continue
            self.open_word_with_cmd(self.doc_helper.tmp_converted_file)
            if not self.errors_handler_when_opening():
                self.close_word_with_cmd()
                continue
            self.events_handler_when_opening()
            self.get_screenshots(StaticData.TMP_DIR_CONVERTED_IMG)
            self.close_word_with_cmd()

            print(f'[bold green]In test[/] {self.doc_helper.source_file}')
            self.open_word_with_cmd(self.doc_helper.tmp_source_file)
            self.get_screenshots(StaticData.TMP_DIR_SOURCE_IMG)
            self.close_word_with_cmd()

            CompareImage()
            self.doc_helper.tmp_cleaner()
