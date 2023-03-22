# -*- coding: utf-8 -*-
from os.path import basename
from rich import print

import settings
from frameworks.StaticData import StaticData
from frameworks.host_control import FileUtils
from .handlers.libre_errors import LibreErrors
from .handlers.libre_events import LibreEvents
from .libre_office import LibreOffice
from settings import version
from loguru import logger

from ..windows_handler.windows import Window


class OpenerLibre:
    extensions = ('.ods', '.odp', '.odt')

    def __init__(self, extension):
        self.extension = extension
        self.converted_file = ''
        self.libre = LibreOffice()
        self.tmp_dir = StaticData.TMP_DIR_IN_TEST
        self.errors = LibreErrors()
        self.events = LibreEvents()
        self.window_control = Window()
        self.wait_for_opening = settings.wait_for_opening
        FileUtils.terminate_process(StaticData.TERMINATE_PROCESS_LIST)

    def open_file(self, file_path: str) -> None:
        if not file_path.lower().endswith(self.extensions):
            return print(f"[bold red]|WARNING| Invalid file extension: {file_path}, use extensions: {self.extensions}")
        print(f'[bold green]In test[/] {basename(file_path)}')
        tmp_file = FileUtils.make_tmp_file(file_path, self.tmp_dir)
        self.libre.open(tmp_file)
        hwnd = self.window_control.wait_until_open(['SALFRAME', 'SALSUBFRAME'], self.wait_for_opening)
        self.window_control.set_size(hwnd, z=1920, w=1080)
        self.events.when_opening()
        self.errors.when_opening()
        self.window_control.close(hwnd)
        self.events.when_closing()
        FileUtils.delete(tmp_file, silence=True)

    def open_files(self, paths_list: list) -> None:
        logger.info(f'Opener {self.extension} with LibreOffice on version: {version} is running.')
        for file_path in paths_list:
            self.open_file(file_path)
