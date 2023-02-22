# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from framework.word import Word
from settings import version


class OpenerDocx(Word):
    def run_opener(self, list_of_files):
        self.doc_helper.terminate_process()
        logger.info(f'Opener {self.doc_helper.converted_extension} with ms Word on version: {version} is running.')
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.lower().endswith(".docx"):
                continue
            self.doc_helper.preparing_files_for_opening_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            self.open_word_with_cmd(self.doc_helper.tmp_converted_file)
            self.events_handler_when_opening(opener_test=True)
            self.errors_handler_when_opening()
            self.close_word_with_cmd()
            self.doc_helper.tmp_cleaner()
