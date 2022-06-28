# -*- coding: utf-8 -*-
from loguru import logger
from rich import print

from config import version
from framework.power_point import PowerPoint
from libs.helpers.helper import Helper


class OpenerPptx:

    def __init__(self, source_extension):
        self.helper = Helper(source_extension, 'pptx')
        self.power_point = PowerPoint(self.helper)
        self.helper.create_dir(self.helper.opener_errors)
        logger.info(f'Opener {self.helper.converted_extension} with ms PowerPoint on version: {version} is running.')

    def run_opener(self, list_of_files):
        for self.helper.converted_file in list_of_files:
            if not self.helper.converted_file.endswith((".pptx", ".PPTX")):
                continue
            self.helper.preparing_files_for_test()
            print(f'[bold green]In test[/] {self.helper.converted_file}')
            self.power_point.open_presentation_with_cmd(self.helper.tmp_name_converted_file)
            self.power_point.errors_handler_when_opening()
            self.power_point.close_presentation_with_hotkey()
            self.helper.tmp_cleaner()
