# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import join
from rich import print

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from frameworks.host_control import FileUtils, HostInfo
from frameworks.onlyoffice.handlers.VersionHandler import VersionHandler
from frameworks.report_action.report import Report
import settings


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
        self.reports_dir = StaticData.reports_dir()
        self.tmp_dir = StaticData.TMP_DIR
        self.os = HostInfo().os
        self.errors_only = settings.errors_only

    # @param [String] input_format
    # @param [String] output_format
    # @param [String] x2t_version
    # @return [String] path to report.csv and tmp report directory.
    def random_report_path(self, x2t_version) -> str:
        random_tmp_dir = FileUtils.random_name(self.tmp_dir)
        FileUtils.create_dir(random_tmp_dir)
        return join(random_tmp_dir, f"{x2t_version}.csv")

    def report_path(self, x2t_version) -> str:
        dir_path = join(self.reports_dir, VersionHandler(x2t_version).without_build, self.os, f"conversion")
        report_path = join(dir_path, f"{x2t_version}.csv")
        FileUtils.create_dir(dir_path, silence=True)
        return report_path

    def merge_x2ttester_reports(self, x2ttester_reports, x2t_version):
        return self.merge_reports(x2ttester_reports, self.report_path(x2t_version))

    @staticmethod
    def print_results(df_errors, errors_array, passed_array, x2ttester_report_csv):
        print(
            f"[bold red]\n{'-' * 90}\n{df_errors}\n{'-' * 90}\nFiles with errors:\n{errors_array}\n"
            f"[bold green]\n{'-' * 90}\nNum passed files:\n{len(passed_array)}\n\n"
            f"[blue]Report path: {x2ttester_report_csv}"
        )

    def report_handler(self, x2ttester_report_csv):
        df = self.pandas_read(x2ttester_report_csv)
        df.rename(columns=self.__columns_names, inplace=True)
        df = df.drop(columns=['Log', 'Input_size', 'Time', 'Output_file'], axis=1)
        df_errors = df[df.Output_size == 0.0] if self.errors_only != '1' else df
        errors_array = [file for file in df_errors.Input_file.unique() if 'Time: ' not in file]
        passed_array = [file for file in df[df.Output_size != 0.0].Input_file.unique() if 'Time: ' not in file]
        self.print_results(df_errors, errors_array, passed_array, x2ttester_report_csv)
