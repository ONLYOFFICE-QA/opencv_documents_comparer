# -*- coding: utf-8 -*-

from loguru import logger
from rich import print

from libs.functional.documents.word import Word
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper


class RtfDocxCompareImg:

    def __init__(self):
        self.source_extension = 'rtf'
        self.converted_extension = 'docx'
        self.helper = Helper(self.source_extension, self.converted_extension)
        self.word = Word(self.helper)

    def run_compare_rtf_docx(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if self.helper.converted_file.endswith((".docx", ".DOCX")):
                self.helper.preparing_files_for_test()

                print(f'[bold green]In test[/bold green] {self.helper.converted_file}')
                # Getting Statistics
                self.word.word_opener(self.helper.tmp_name)

                if self.word.statistics_word is not None:
                    print(f"[bold blue]Number of pages:[/bold blue] {self.word.statistics_word['num_of_sheets']}")
                    self.word.open_word_with_cmd(self.helper.tmp_name_converted_file)
                    self.word.get_screenshots(self.helper.tmp_dir_converted_image)
                    self.word.close_word_with_cmd()

                    print(f'[bold green]In test[/bold green] {self.helper.source_file}')
                    self.word.open_word_with_cmd(self.helper.tmp_name_source_file)
                    self.word.get_screenshots(self.helper.tmp_dir_source_image)
                    self.word.close_word_with_cmd()

                    CompareImage(self.helper)

                elif self.word.statistics_word is None:
                    logger.error(f"Can't get number of pages in {self.helper.source_file}. Copied files "
                                 "to 'failed_to_open_file'")
                    self.helper.copy_to_folder(self.helper.failed_source)
                else:
                    logger.debug(f"Debug. Statistics word: {self.word.statistics_word}")

            self.helper.tmp_cleaner()
