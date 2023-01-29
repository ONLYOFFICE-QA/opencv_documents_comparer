# -*- coding: utf-8 -*-
import csv

import pandas as pd
from rich import print

from framework.singleton import singleton


@singleton
class ReportActions:
    def __init__(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)

    @staticmethod
    def csv_writer(file_path, mode, message):
        with open(file_path, mode, newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter='\t')
            writer.writerow(message)

    @staticmethod
    def out_x2ttester_report_csv(x2ttester_report_csv):
        df = pd.read_csv(x2ttester_report_csv, delimiter="\t")
        df.rename(
            columns={
                'Input file': 'Input_file', 'Output file': 'Output_file', 'Input size': 'Input_size',
                'Output size': 'Output_size', 'Exit code': 'Exit_code', 'Log': 'Log', 'Time': 'Time'
            }, inplace=True
        )
        df = df.drop(columns=['Log', 'Input_size', 'Time', 'Output_file'], axis=1)
        df_errors = df[df.Output_size == 0.0]
        errors_array = [file for file in df_errors.Input_file.unique() if 'Time: ' not in file]
        passed_array = [file for file in df[df.Output_size != 0.0].Input_file.unique() if 'Time: ' not in file]
        print(f"[bold red]\n{'-' * 90}\n{df_errors}\n{'-' * 90}\nFiles with errors:\n{errors_array}")
        print(f"[bold green]\n{'-' * 90}\nNum passed files:\n{len(passed_array)}")
        print(f"[blue]Report path: {x2ttester_report_csv}")
