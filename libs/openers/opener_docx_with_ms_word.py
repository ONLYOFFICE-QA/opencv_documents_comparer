# -*- coding: utf-8 -*-
from config import version
from framework.word import Word
from management import *


class OpenerDocx(Word):
    def run_opener(self, list_of_files):
        logger.info(f'Opener {self.doc_helper.converted_extension} with ms Word on version: {version} is running.')
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.endswith((".docx", ".DOCX")):
                continue
            self.doc_helper.preparing_files_for_opening_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            self.open_word_with_cmd(self.doc_helper.tmp_converted_file)
            self.errors_handler_when_opening()
            self.close_word_with_cmd()
            self.doc_helper.tmp_cleaner()
