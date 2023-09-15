# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import join, dirname, realpath
from rich import print

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from host_control import FileUtils, HostInfo
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
        self.exceptions = FileUtils.read_json(f"{dirname(realpath(__file__))}/../assets/conversion_exception.json")
        self.reports_dir = StaticData.reports_dir()
        self.tmp_dir = StaticData.tmp_dir
        self.errors_only: bool = config.errors_only
        self.x2t_dir = StaticData.core_dir()

    # @param [String] input_format
    # @param [String] output_format
    # @param [String] x2t_version
    # @return [String] path to report.csv and tmp report directory.
    @staticmethod
    def tmp_file(dir_path: str) -> str:
        random_tmp_dir = FileUtils.random_name(dir_path)
        FileUtils.create_dir(random_tmp_dir, stdout=False)
        return join(random_tmp_dir, FileUtils.random_name(random_tmp_dir, extension='.csv'))

    def path(self, x2t_version: str) -> str:
        dir_path = join(self.reports_dir, VersionHandler(x2t_version).without_build, "conversion", HostInfo().os)
        FileUtils.create_dir(dir_path, stdout=False)
        return join(dir_path, f"{x2t_version}_{datetime.now().strftime('%H_%M_%S')}.csv")

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
        self._send_to_telegram([processed_report, report_path], tg_msg) if tg_msg else ...
        self._print_results(df, errors_list, passed_num, report_path)

    def _errors_list(self, df) -> list:
        errors = df[df.Output_size == 0.0] if not self.errors_only else df
        return [file for file in errors[df.BugInfo == 0].Input_file.unique()]

    def _bug_info(self, row) -> str | int:
        for i in self.exceptions.items():
            if row.Input_file in i[1]['files']:
                if HostInfo().os in i[1]['os'] or not i[1]['os']:
                    if row.Direction in i[1]['directions'] or not i[1]['directions']:
                        return f"{i[1]['description']} {i[1]['link']}" if i[1]['link'] or i[1]['description'] else '1'
        return 0

    @staticmethod
    def _add_to_end(df, column_name: str, value: str | int | float):
        df.loc[df.index.max() + 1, column_name] = value

    @staticmethod
    def _send_to_telegram(reports: list, tg_msg: str) -> None:
        Telegram().send_media_group(reports, caption=tg_msg)
