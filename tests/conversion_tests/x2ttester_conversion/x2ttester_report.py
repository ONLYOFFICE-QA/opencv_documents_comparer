# -*- coding: utf-8 -*-
import json
import os
import zipfile
from datetime import datetime
from os.path import isdir, join, splitext
from re import sub

import pandas as pd
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
    """
    X2ttesterReport class handles the generation and management of reports for X2T tester conversions.
    It provides functionalities to merge reports, handle exceptions, and output results.
    """

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
        """
        Initializes the X2ttesterReport with the given test configuration and exceptions.

        :param test_config: Configuration for the X2T tester.
        :param exceptions_json: JSON object containing exception details.
        """
        super().__init__()
        self.config = test_config
        self.reports_dir = self.config.reports_dir
        self.exceptions = exceptions_json
        self.os = HostInfo().os
        self._errors_xml = set()
        self._errors_xml_found = None

    def collect_errors_xml(self, output_dir: str) -> None:
        """
        Remembers failed conversions of an x2ttester run by the parameters xml that x2ttester saves to _errors.

        :param output_dir: Output directory of the x2ttester run.
        """
        errors_dir = join(output_dir, '_errors')
        found = isdir(errors_dir)
        if found:
            self._errors_xml.update(os.listdir(errors_dir))
        self._errors_xml_found = found if self._errors_xml_found is None else self._errors_xml_found and found

    def path(self) -> str:
        """
        Constructs and returns the path for the report based on the current configuration and time.

        :return: The path where the report will be saved.
        """
        version = VersionHandler(self.config.x2t_version).without_build if self.config.x2t_version else 'unknown'
        dir_path = join(self.reports_dir, version, "conversion", self.os)
        Dir.create(dir_path, stdout=False)
        return join(dir_path, f"{self.config.x2t_version}_{datetime.now().strftime('%H_%M_%S')}.csv")

    def merge_reports(self, x2ttester_report: list) -> str | None:
        """
        Merges multiple X2T tester reports into a single report.

        :param x2ttester_report: List of reports to merge.
        :return: The path to the merged report or None if merging fails.
        """
        return self.merge(list(filter(lambda x: x is not None, x2ttester_report)), self.path())

    @staticmethod
    def _print_results(df, errors: list, passed_num: int | str, report_path: str) -> None:
        """
        Prints the results of the conversion tests including errors and passed files.

        :param df: DataFrame containing the test results.
        :param errors: List of errors encountered during the tests.
        :param passed_num: Number of passed files.
        :param report_path: Path to the report file.
        """
        print(
            f"[bold red]\n{'-' * 90}\n{df,}\n{'-' * 90}\n"
            f"Files with errors:\n{errors}\n"
            f"[bold green]\n{'-' * 90}\nNum passed files:\n{len(passed_num)}\n\n"
            f"[cyan]\n{'-' * 90}\nReport path: {report_path}\n{'-' * 90}\n"
        )

    def handler(self, report_path: str, tg_msg: str = None) -> None:
        """
        Processes the report, updates it with bug information, and optionally sends a Telegram message.

        :param report_path: Path to the report file to process.
        :param tg_msg: Optional Telegram message to send.
        """
        df = self.read(report_path)
        df.rename(columns=self.__columns_names, inplace=True)
        df = df.drop(columns=['Log', 'Input_size', 'Output_file'])
        df.insert(df.columns.get_loc('Direction') + 1, 'BugInfo', df.apply(self._bug_info, axis=1))
        df = df[~df['Input_file'].str.contains('Time: ')]

        errors_list = self._errors_list(df)
        passed_num = f"{len([file for file in df[df.Output_size != 0.0].Input_file.unique()])}"

        # With errors_only disabled x2ttester reports every conversion: keep only failures in the errors report
        # and send all rows as a separate compressed full report.
        errors_df = df if self.config.errors_only else self._failed_conversions(df).reset_index(drop=True)
        full_report = report_path if self.config.errors_only else self._full_report(df)

        self._add_to_end(errors_df, 'BugInfo', f"Errors: {errors_list}")
        self._add_to_end(errors_df, 'BugInfo', f"Passed: {passed_num}")

        processed_report = self.save_csv(errors_df, self.path())
        self._print_results(errors_df, errors_list, passed_num, report_path)

        if tg_msg:
            self._send_to_telegram(
                [
                    self._rename_report_for_tg(processed_report, f'{self.config.x2t_version}_errors_only.csv'),
                    self._rename_report_for_tg(
                        full_report,
                        f'{self.config.x2t_version}_full{splitext(full_report)[1]}'
                    ),
                ],
                f"{tg_msg}\n\nStatus: `{'Some files have errors' if errors_list else 'All tests passed'}`"
            )

    def _full_report(self, df) -> str:
        """
        Saves all conversions (without the parameters xml from the Log column) and compresses the report:
        a full run has hundreds of thousands of rows, too big for a Telegram document as plain csv.

        :param df: DataFrame with all conversions.
        :return: Path to the zip archive with the full report.
        """
        csv_path = self.save_csv(df, f"{splitext(self.path())[0]}_full.csv")
        zip_path = f"{splitext(csv_path)[0]}.zip"
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as archive:
            archive.write(csv_path, f"{self.config.x2t_version}_full.csv")
        return zip_path

    def _rename_report_for_tg(self, report_path: str, new_name: str) -> str:
        """
        Renames the report for Telegram by copying it to a temporary directory.

        :param report_path: Original path to the report.
        :param new_name: New name for the report.
        :return: Path to the renamed report.
        """
        new_path = join(self.config.tmp_dir, new_name)
        File.delete(new_path, stdout=False, stderr=False)
        File.copy(report_path, new_path, stdout=False)
        return new_path

    def _errors_list(self, df) -> list:
        """
        Extracts the list of errors from the given DataFrame.

        :param df: DataFrame containing the test results.
        :return: List of errors encountered during the tests.
        """
        if df.empty:
            return []

        errors = self._failed_conversions(df) if not self.config.errors_only else df
        mask = (errors.BugInfo == 0)
        return errors.loc[mask, 'Input_file'].unique().tolist()

    def _failed_conversions(self, df) -> pd.DataFrame:
        """
        Selects failed conversions from a report with all conversions, the same way x2ttester does in errorsOnly
        mode.

        :param df: DataFrame with all conversions.
        :return: DataFrame with failed conversions only.
        """
        if not self._errors_xml_found:  # x2ttester without the _errors folder (before 7.3)
            return df[pd.to_numeric(df.Exit_code, errors='coerce').ne(0)]

        output_ext = df.Direction.astype(str).str.split('-').str[-1]
        return df[(df.Input_file.astype(str) + '_.' + output_ext + '.xml').isin(self._errors_xml)]

    def _bug_info(self, row) -> str | int:
        """
        Extracts bug information for the given row.

        :param row: Row from the DataFrame containing the test results.
        :return: Bug information as a string or integer.
        """
        for exception in self.exceptions.values():
            if not self._file_math(row.Input_file, exception['files']):
                continue

            if self._os_math(exception['os']) and self._direction_match(row.Direction, exception['directions']):
                description = exception.get('description', '')
                link = exception.get('link', '')

                if description or link:
                    return f"{description} {link}".strip()

                return '1'

        return 0

    @staticmethod
    def _file_math(input_file: str, exception_files: list) -> bool:
        """
        Checks if the input file matches the exception files.

        :param input_file: Input file to check.
        :param exception_files: List of exception files.
        :return: True if the input file matches, False otherwise.
        """
        if input_file in exception_files:
            return True

        if exception_files and exception_files[0].startswith('*.'):
            return splitext(input_file)[1] == splitext(exception_files[0])[1]

        return False

    def _os_math(self, exception_os: list) -> bool:
        """
        Checks if the current OS matches the exception OS.

        :param exception_os: List of exception OS.
        :return: True if the current OS matches, False otherwise.
        """
        if not exception_os:
            return True

        return self.os in exception_os

    @staticmethod
    def _direction_match(input_direction: str, exception_directions: list) -> bool:
        """
        Checks if the input direction matches the exception directions.

        :param input_direction: Input direction to check.
        :param exception_directions: List of exception directions.
        :return: True if the input direction matches, False otherwise.
        """
        if not exception_directions:
            return True

        if input_direction in exception_directions:
            return True

        for exc in exception_directions:
            if exc.startswith("*-"):
                return sub(r'^.*(-png)$', r'*\1', input_direction) == exc

        return False

    @staticmethod
    def _add_to_end(df, column_name: str, value: str | int | float):
        """
        Adds a new row to the end of the DataFrame with the given value.

        :param df: DataFrame to update.
        :param column_name: Name of the column to update.
        :param value: Value to add to the column.
        """
        df[column_name] = df[column_name].astype(str)
        df.loc[len(df.index), column_name] = value

    @staticmethod
    def _send_to_telegram(reports: list, tg_msg: str) -> None:
        """
        Sends a Telegram message with the given reports.

        :param reports: List of reports to send.
        :param tg_msg: Telegram message to send.
        """
        Telegram().send_media_group(reports, caption=tg_msg)
