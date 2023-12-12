# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import join, basename, dirname, realpath

from host_tools import Dir, File
from rich import print
from re import sub, search
import pandas as pd

import config
from frameworks.decorators import singleton
from frameworks.report import Report


@singleton
class CompareReport:
    def __init__(self, reports_dir: str):
        self.exceptions = File.read_json(f"{dirname(realpath(__file__))}/../assets/opener_exception.json")
        self.time_pattern = f"{datetime.now().strftime('%H_%M_%S')}"
        self.report_dir = reports_dir
        self.path = join(self.report_dir, f"{config.version}_compare_{self.time_pattern}.csv")
        self._set_pandas_options()
        Dir.create(self.report_dir, stdout=False)
        Report.write(self.path, 'w', ['File_name', 'Direction', 'Exit_code', 'Bug_info', 'Version'])

    def write(self, file_path: str, exit_code: int | str) -> None:
        file_name, dir_name = basename(file_path), basename(dirname(file_path))
        direction = sub(r'(\d+)_(\w+)_(\w+)', r'\2-\3', dir_name.split('.')[-1])
        find_version = search(r"\d+\.\d+\.\d+\.\d+", dir_name)
        version = find_version.group() if find_version else 0
        for i in self.exceptions.items():
            if direction in i[1]['directions'] and file_name in i[1]['files']:
                bug_info = i[1]['link'] if i[1]['link'] else i[1]['description'] if i[1]['description'] else 1
                return Report.write(self.path, 'a', [file_name, direction, exit_code, bug_info, version])
        Report.write(self.path, 'a', [file_name, direction, exit_code, 0, version])

    def handler(self):
        print(f"[bold red]{Report.read(self.path)}")
        print(f"\n\n[bold cyan]{'-' * 90}\nPath to opener report: {self.path}\n{'-' * 90}")

    @staticmethod
    def _set_pandas_options():
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)
