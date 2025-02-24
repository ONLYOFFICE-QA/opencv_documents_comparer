# -*- coding: utf-8 -*-
import json
import os
from datetime import datetime
from os.path import join, splitext
from re import sub

from host_tools.utils import Dir
from rich import print

from frameworks.decorators import singleton
from host_tools import File, HostInfo
from frameworks.editors.onlyoffice.handlers.VersionHandler import VersionHandler
from frameworks.report import Report

from telegram import Telegram

from .x2ttester_test_config import X2ttesterTestConfig



@singleton
class X2ttesterReport(Report):
    __columns_names = {
        'Input file': 'Input_file',
        'Output file': 'Output_file',
        'Input size': 'Input_size',
        'Output size': 'Output_size',
        'Exit code': 'Exit_code',
        'Log': 'Log',
        'Time': 'Time'
    }

    def __init__(self, test_config: X2ttesterTestConfig, exceptions_json: json):
        super().__init__()
        self.config = test_config
        self.reports_dir = self.config.reports_dir
        self.exceptions = exceptions_json
        self.os = HostInfo().os

    def path(self) -> str:
        dir_path = join(self.reports_dir, VersionHandler(self.config.x2t_version).without_build, "conversion", self.os)
        Dir.create(dir_path, stdout=False)
        return join(dir_path, f"{self.config.x2t_version}_{datetime.now().strftime('%H_%M_%S')}.csv")

    def merge_reports(self, x2ttester_report: list) -> str | None:
        return self.merge(list(filter(lambda x: x is not None, x2ttester_report)), self.path())

    @staticmethod
    def _print_results(df, errors: list, passed_num: int | str, report_path: str) -> None:
        print(
            f"[bold red]\n{'-' * 90}\n{df,}\n{'-' * 90}\n"
            f"Files with errors:\n{errors}\n"
            f"[bold green]\n{'-' * 90}\nNum passed files:\n{len(passed_num)}\n\n"
            f"[cyan]\n{'-' * 90}\nReport path: {report_path}\n{'-' * 90}\n"
        )

    def handler(self, report_path: str, tg_msg: str = None) -> None:
        df = self.read(report_path)
        df.rename(columns=self.__columns_names, inplace=True)
        df = df.drop(columns=['Log', 'Input_size', 'Output_file'], axis=1)
        df = df[~df['Input_file'].str.contains('Time: ')]

        df.insert(df.columns.get_loc('Direction') + 1, 'BugInfo', df.apply(self._bug_info, axis=1))

        errors_list = self._errors_list(df)
        passed_num = f"{len([file for file in df[df.Output_size != 0.0].Input_file.unique()])}"

        self._add_to_end(df, 'BugInfo', f"Errors: {errors_list}")
        self._add_to_end(df, 'BugInfo', f"Passed: {passed_num}")

        processed_report = self.save_csv(df, self.path())
        self._print_results(df, errors_list, passed_num, report_path)

        if tg_msg:
            self._send_to_telegram(
                [
                    self._rename_report_for_tg(processed_report, f'{self.config.x2t_version}_errors_only.csv'),
                    self._rename_report_for_tg(report_path, f'{self.config.x2t_version}_full.csv'),
                ],
                f"{tg_msg}\n\nStatus: `{'Some files have errors' if errors_list else 'All tests passed'}`"
            )

    def _rename_report_for_tg(self, report_path: str, new_name: str) -> str:
        new_path = join(self.config.tmp_dir, new_name)
        File.delete(new_path, stdout=False, stderr=False)
        File.copy(report_path, new_path, stdout=False)
        return new_path

    def _errors_list(self, df) -> list:
        errors = df[df.Output_size == 0.0] if not self.config.errors_only else df
        mask = (errors.BugInfo == 0)
        return errors.loc[mask, 'Input_file'].unique().tolist()

    def _bug_info(self, row) -> str | int:
        for exception in self.exceptions.values():
            if not self._file_math(row.Input_file, exception['files']):
                continue

            if self._os_math(exception.get('os')) and self._direction_match(row.Direction, exception['directions']):
                description = exception.get('description', '')
                link = exception.get('link', '')

                if description or link:
                    return f"{description} {link}".strip()

                return '1'

        return 0

    @staticmethod
    def _file_math(input_file: str, exception_files: list) -> bool:
        return(
                input_file in exception_files or
                (
                        exception_files and
                        exception_files[0].startswith('*.') and
                        splitext(exception_files[0])[1] == splitext(input_file)[1]
                )
        )
    
    def _os_math(self, exception_os: list) -> bool:
        return (not exception_os) or (self.os in exception_os)
    
    @staticmethod
    def _direction_match(input_directions: str, exception_directions: list):
        return (
                (not exception_directions) or
                (input_directions in exception_directions) or
                (
                    exception_directions.startswith('*-') and
                    sub(r'^.*(-png)$', r'*\1', input_directions) in exception_directions
                )
        )

    @staticmethod
    def _add_to_end(df, column_name: str, value: str | int | float):
        df[column_name] = df[column_name].astype(str)
        df.loc[len(df.index), column_name] = value

    @staticmethod
    def _send_to_telegram(reports: list, tg_msg: str) -> None:
        Telegram().send_media_group(reports, caption=tg_msg)
