# -*- coding: utf-8 -*-
from os.path import join, basename, dirname, realpath, exists, splitext, isfile
from rich import print
from re import sub, search
import pandas as pd

from frameworks.host_control import FileUtils
from frameworks.report import Report
from frameworks.telegram import Telegram


class OpenerReport:
    def __init__(self, reports_path: str):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)
        self.exceptions = FileUtils.read_json(f"{dirname(realpath(__file__))}/../assets/opener_exception.json")
        self.path = reports_path
        self.errors_path = f"{splitext(self.path)[0]}(errors_only).csv"
        FileUtils.create_dir(dirname(self.path), silence=True)

    def _write_titles(self):
        self._writer('w', ['File_name', 'Direction', 'Exit_code', 'Bug_info', 'Version', 'File_path'])

    def write(self, file_path: str, exit_code: int | str) -> None:
        self._write_titles() if not isfile(self.path) else ...
        name, dir_name = basename(file_path), basename(dirname(file_path))
        direction = self._direction(dir_name)
        self._writer(
            'a', [name, direction, exit_code, self._bug_info(direction, name), self._version(dir_name), file_path]
        )

    def handler(self, tg_msg: bool | str = False) -> None:
        df = Report.read(self.path)
        df_errors = df[(df.Exit_code != '0') | (df.Bug_info != '0')]
        error_files = [file for file in df[(df.Exit_code != '0') & (df.Bug_info == '0')].File_name]
        print(f"[bold red]{df_errors.drop(columns=['File_path'], axis=1)}\n\nErrors files: {error_files}\n")
        self._add_to_end(df_errors, 'File_path', f"Errors files: {error_files}")
        errors_report = Report.save_csv(df_errors, self.errors_path)
        Telegram().send_media_group([errors_report, self.path], caption=tg_msg) if tg_msg else ...
        print(f"\n\n[bold cyan]{'-' * 90}\nPath to opener report: {self.path}\n"
              f"Errors_report: {errors_report}\n{'-' * 90}")

    @staticmethod
    def _add_to_end(df, column_name: str, value: str | int | float):
        df.loc[df.index.max() + 1, column_name] = value

    def _bug_info(self, direction: str, file_name: str) -> int | str:
        for i in self.exceptions.items():
            if direction in i[1]['directions'] and file_name in i[1]['files']:
                return i[1]['link'] if i[1]['link'] else i[1]['description'] if i[1]['description'] else '1'
        return 0

    def tested_files(self) -> list:
        if exists(self.path):
            return [join(basename(dirname(path)), basename(path)) for path in Report.read(self.path).File_path]
        return []

    def _writer(self, mode: str, info: list) -> None:
        Report.write(self.path, mode, info)

    @staticmethod
    def _direction(dir_name: str) -> str:
        return sub(r'(\d+)_(\w+)_(\w+)', r'\2-\3', dir_name.split('.')[-1])

    @staticmethod
    def _version(dir_name: str) -> str:
        find_version = search(r"\d+\.\d+\.\d+\.\d+", dir_name)
        return find_version.group() if find_version else 0
