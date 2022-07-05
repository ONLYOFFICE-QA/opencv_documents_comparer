# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from config import version
from framework.word import Word


class OpenerDocx:
    def __init__(self, source_extension):
        self.word = Word(source_extension, 'docx')
        self.helper = self.word.helper
        self.helper.create_dir(self.helper.opener_errors)
        logger.info(f'Opener {self.helper.converted_extension} with ms Word on version: {version} is running.')

    def run_opener_word(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".docx", ".DOCX")):
                continue
            self.helper.preparing_files_for_test()

            print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
            self.word.open_word_with_cmd(self.helper.tmp_name_converted_file)
            self.word.errors_handler_when_opening()
            self.word.close_word_with_cmd()
            self.helper.tmp_cleaner()
