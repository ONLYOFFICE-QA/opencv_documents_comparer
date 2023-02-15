# -*- coding: utf-8 -*-
import csv
from os.path import dirname

import pandas as pd
from rich import print

import settings
from framework.FileUtils import FileUtils
from framework.singleton import singleton


@singleton
class ReportActions:
    def __init__(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)

    @staticmethod
    def merge_reports(reports_array, path_to_result_csv):
        df = pd.concat([pd.read_csv(csv_) for csv_ in reports_array], ignore_index=True)
        df.to_csv(path_to_result_csv, index=False)

    @staticmethod
    def csv_writer(file_path, mode, message):
        FileUtils.create_dir(dirname(file_path))
        with open(file_path, mode, newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter='\t')
            writer.writerow(message)

    @staticmethod
    def read_csv_via_pandas(csv_file, delimiter="\t"):
        return pd.read_csv(csv_file, delimiter=delimiter)

    def out_x2ttester_report_csv(self, x2ttester_report_csv):
        df = self.read_csv_via_pandas(x2ttester_report_csv)
        df.rename(
            columns={
                'Input file': 'Input_file', 'Output file': 'Output_file', 'Input size': 'Input_size',
                'Output size': 'Output_size', 'Exit code': 'Exit_code', 'Log': 'Log', 'Time': 'Time'
            }, inplace=True
        )
        df = df.drop(columns=['Log', 'Input_size', 'Time', 'Output_file'], axis=1)
        df_errors = df[df.Output_size == 0.0] if settings.errors_only != '1' else df
        errors_array = [file for file in df_errors.Input_file.unique() if 'Time: ' not in file]
        passed_array = [file for file in df[df.Output_size != 0.0].Input_file.unique() if 'Time: ' not in file]
        print(f"[bold red]\n{'-' * 90}\n{df_errors}\n{'-' * 90}\nFiles with errors:\n{errors_array}\n"
              f"[bold green]\n{'-' * 90}\nNum passed files:\n{len(passed_array)}\n\n"
              f"[blue]Report path: {x2ttester_report_csv}")
