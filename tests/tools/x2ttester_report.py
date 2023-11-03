# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import join, dirname, realpath

from host_control.utils import Dir
from rich import print

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from host_control import File, HostInfo
from frameworks.editors.onlyoffice.handlers.VersionHandler import VersionHandler
from frameworks.report import Report
import config
from frameworks.telegram import Telegram


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

    def __init__(self):
        super().__init__()
        self.exceptions = File.read_json(f"{dirname(realpath(__file__))}/../assets/conversion_exception.json")
        self.reports_dir = StaticData.reports_dir()
        self.tmp_dir = StaticData.tmp_dir
        self.errors_only: bool = config.errors_only
        self.x2t_dir = StaticData.core_dir()
        self.os = HostInfo().os


    def path(self, x2t_version: str) -> str:
        dir_path = join(self.reports_dir, VersionHandler(x2t_version).without_build, "conversion", HostInfo().os)
        Dir.create(dir_path, stdout=False)
        return join(dir_path, f"{x2t_version}_{datetime.now().strftime('%H_%M_%S')}.csv")

    def tmp_file(self):
        tmp_report = File.unique_name(File.unique_name(self.tmp_dir), 'csv')
        Dir.create(dirname(tmp_report), stdout=False)
        return tmp_report

    def merge_reports(self, x2ttester_report: list, x2t_version: str) -> str | None:
        return self.merge(list(filter(lambda x: x is not None, x2ttester_report)), self.path(x2t_version))

    @staticmethod
    def _print_results(df, errors: list, passed_num: int | str, report_path: str) -> None:
        print(
            f"[bold red]\n{'-' * 90}\n{df,}\n{'-' * 90}\n"
            f"Files with errors:\n{errors}\n"
            f"[bold green]\n{'-' * 90}\nNum passed files:\n{len(passed_num)}\n\n"
            f"[cyan]\n{'-' * 90}\nReport path: {report_path}\n{'-' * 90}\n"
        )

    def handler(self, report_path: str, x2t_version: str, tg_msg: str = None) -> None:
        df = self.read(report_path)
        df.rename(columns=self.__columns_names, inplace=True)
        df = df.drop(columns=['Log', 'Input_size', 'Output_file'], axis=1)
        df = df[~df['Input_file'].str.contains('Time: ')]

        df.insert(df.columns.get_loc('Direction') + 1, 'BugInfo', df.apply(self._bug_info, axis=1))

        errors_list = self._errors_list(df)
        passed_num = f"{len([file for file in df[df.Output_size != 0.0].Input_file.unique()])}"

        self._add_to_end(df, 'BugInfo', f"Errors: {errors_list}")
        self._add_to_end(df, 'BugInfo', f"Passed: {passed_num}")

        processed_report = self.save_csv(df, self.path(x2t_version))
        self._print_results(df, errors_list, passed_num, report_path)

        if tg_msg:
            self._send_to_telegram(
                [
                    self._rename_report_for_tg(processed_report, f'{x2t_version}_errors_only.csv'),
                    self._rename_report_for_tg(report_path, f'{x2t_version}_full.csv'),
                ],
                f"{tg_msg}\n\nStatus: `{'Some files have errors' if errors_list else 'All tests passed'}`"
            )

    def _rename_report_for_tg(self, report_path: str, new_name: str) -> str:
        new_path = join(self.tmp_dir, new_name)
        File.delete(new_path, stdout=False)
        File.copy(report_path, new_path)
        return new_path

    def _errors_list(self, df) -> list:
        errors = df[df.Output_size == 0.0] if not self.errors_only else df
        return [file for file in errors[df.BugInfo == 0].Input_file.unique()]

    def _bug_info(self, row) -> str | int:
        for _, value in self.exceptions.items():
            if row.Input_file in value['files']:
                if self.os in value['os'] or not value['os']:
                    if row.Direction in value['directions'] or not value['directions']:
                        if value['link'] or value['description']:
                            return f"{value['description']} {value['link']}".strip()
                        return '1'
        return 0

    @staticmethod
    def _add_to_end(df, column_name: str, value: str | int | float):
        df.loc[len(df.index), column_name] = value

    @staticmethod
    def _send_to_telegram(reports: list, tg_msg: str) -> None:
        Telegram().send_media_group(reports, caption=tg_msg)
