# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from config import version
from framework.word import Word
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper


class DocDocxCompareImg:
    def __init__(self):
        self.helper = Helper('doc', 'docx')
        self.word = Word(self.helper)
        logger.info(f'The {self.helper.source_extension} to {self.helper.converted_extension} '
                    f'comparison on version: {version} is running.')

    def run_compare_word(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".docx", ".DOCX")):
                continue
            self.helper.preparing_files_for_test()

            print(f'[bold green]In test[/] {self.helper.converted_file}')
            if not self.word.word_opener(self.helper.tmp_name):  # Getting Statistics
                continue
            self.word.open_word_with_cmd(self.helper.tmp_name_converted_file)
            if not self.word.errors_handler_when_opening():
                self.word.close_word_with_cmd()
                continue
            self.word.events_handler_when_opening()
            self.word.get_screenshots(self.helper.tmp_dir_converted_image)
            self.word.close_word_with_cmd()

            print(f'[bold green]In test[/] {self.helper.source_file}')
            self.word.open_word_with_cmd(self.helper.tmp_name_source_file)
            self.word.get_screenshots(self.helper.tmp_dir_source_image)
            self.word.close_word_with_cmd()

            CompareImage(self.helper)
            self.helper.tmp_cleaner()
