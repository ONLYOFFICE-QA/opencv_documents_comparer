# -*- coding: utf-8 -*-
from os.path import join, basename, dirname, realpath, exists, splitext, isfile

from host_tools.utils import Str
from rich import print
from re import sub
import pandas as pd

from host_tools import File, Dir
from frameworks.report import Report
from telegram import Telegram


class OpenerReport:
    titles: list = ['File_name', 'Direction', 'Exit_code', 'Bug_info', 'Version', 'Os', 'Mode', 'File_path']
    os_pattern: str = r'\(os_(.*?)\)'
    mod_pattern: str = r'\(mode_(.*?)\)'
    version_pattern: str = r"(\d+).(\d+).(\d+).(\d+)"
    direction_pattern: str = r"\(dir_(\w+-\w+)\)"

    def __init__(self, reports_path: str):
        self._set_pandas_options()
        self.exceptions = File.read_json(f"{dirname(realpath(__file__))}/../assets/opener_exception.json")
        self.path = reports_path
        self.errors_path = f"{splitext(self.path)[0]}(errors_only).csv"
        Dir.create(dirname(self.path), stdout=False)

    def write(self, file_path: str, exit_code: int | str) -> None:
        self._write_titles() if not isfile(self.path) else ...

        direction = Str.search(file_path, self.direction_pattern, group_num=1)
        version = Str.search(file_path, self.version_pattern, group_num=0)
        os = Str.search(file_path, self.os_pattern, group_num=1)
        mode = Str.search(file_path, self.mod_pattern, group_num=1)
        source_name = self._get_source_file_name(file_name=basename(file_path), direction=direction)
        bug_info = self._bug_info(direction, source_name, mode)

        self._writer('a', [source_name, direction, exit_code, bug_info, version, os, mode, file_path])


    def handler(self, tg_msg: bool | str = False) -> None:
        df = Report.read(self.path)
        df_errors = df[(df.Exit_code != '0') | (df.Bug_info != '0')]
        error_files = [file for file in df[(df.Exit_code != '0') & (df.Bug_info == '0')].File_name]

        print(f"[bold red]{df_errors.drop(columns=['File_path'], axis=1)}\n\nErrors files: {error_files}\n")
        self._add_to_end(df_errors, 'File_path', f"Errors files: {error_files}")

        errors_report = Report.save_csv(df_errors, self.errors_path)
        print(f"\n[bold cyan]{'-' * 90}\nPath to report: {self.path}\nErrors_report: {errors_report}\n{'-' * 90}")

        if tg_msg:
            Telegram().send_media_group(
                [errors_report, self.path],
                caption=f"{tg_msg}\n\nStatus: `{'Some files have errors' if error_files else 'All tests passed'}`"
            )

    def _get_direction(self, file_path: str) -> str:
        direction = Str.search(file_path, self.direction_pattern)
        return sub(self.direction_pattern, r'\1-\2', Str.search(file_path, self.direction_pattern))

    def tested_files(self) -> list:
        if exists(self.path):
            return [join(basename(dirname(path)), basename(path)) for path in Report.read(self.path).File_path]
        return []

    @staticmethod
    def _add_to_end(df, column_name: str, value: str | int | float):
        df.loc[df.index.max() + 1, column_name] = value

    def _bug_info(self, direction: str, source_file_name: str, mode: str) -> int | str:
        mode = mode or 'Default'
        for _, info in self.exceptions.items():
            if (
                    direction in info['directions']
                    and source_file_name in info['files']
                    and mode in info['mode']
            ):
                return info['link'] if info['link'] else info['description'] if info['description'] else '1'
        return '0'

    def _writer(self, mode: str, info: list) -> None:
        Report.write(self.path, mode, info)

    def _write_titles(self):
        self._writer('w', self.titles)

    @staticmethod
    def _set_pandas_options():
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)

    @staticmethod
    def _get_source_file_name(file_name: str, direction: str) -> str:
        _extension_list = direction.split('-')
        if _extension_list and len(_extension_list) >= 2:
            source_extension = f".{_extension_list[0].replace('.', '')}"
            return file_name.replace(splitext(file_name)[1], source_extension)
        return file_name
