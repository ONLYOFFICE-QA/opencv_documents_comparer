# -*- coding: utf-8 -*-
from loguru import logger
from rich import print
import pyautogui as pg

from config import version
from framework.word import Word
from libs.helpers.helper import Helper


class OpenerDocx:
    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'docx')
        self.word = Word(self.helper)
        logger.info(f'Opener {self.helper.converted_extension} with ms Word on version: {version} is running.')

    def run_opener_word(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".docx", ".DOCX")):
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
                self.word.open_word_with_cmd_for_opener(self.helper.tmp_name_converted_file)
                if self.word.check_errors.errors \
                        and self.word.check_errors.errors[0] == "#32770" \
                        and self.word.check_errors.errors[1] == "Microsoft Word":

                    logger.error(f"'an error has occurred while opening the file'. "
                                 f"Copied files: {self.helper.converted_file} "
                                 f"and {self.helper.source_file} to 'failed_to_open_converted_file'")

                    pg.press('esc', presses=3, interval=0.2)
                    self.word.close_word_with_cmd()
                    self.helper.copy_to_folder(self.helper.untested_folder)
                    self.word.check_errors.errors.clear()
                elif not self.word.check_errors.errors:
                    self.word.close_word_with_cmd()

                else:
                    logger.debug(f"New Error "
                                 f"Error message: {self.word.check_errors.errors} "
                                 f"Filename: {self.helper.converted_file}")
                    self.word.close_word_with_cmd()
                    self.helper.copy_to_folder(self.helper.failed_source)
            self.helper.tmp_cleaner()
