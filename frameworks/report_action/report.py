# -*- coding: utf-8 -*-
import csv
from os.path import dirname

import pandas as pd
from frameworks.host_control.FileUtils import FileUtils


class Report:
    def __init__(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)

    def merge_reports(self, reports_array, result_csv_path, delimiter='\t') -> str:
        df = pd.concat([self.pandas_read(csv_, delimiter) for csv_ in reports_array], ignore_index=True)
        df.to_csv(result_csv_path, index=False, sep=delimiter)
        return result_csv_path

    @staticmethod
    def csv_write(file_path: str, mode: str, message: str) -> None:
        FileUtils.create_dir(dirname(file_path))
        with open(file_path, mode, newline='') as csv_file:
            writer = csv.writer(csv_file, delimiter='\t')
            writer.writerow(message)

    @staticmethod
    def pandas_read(csv_file, delimiter="\t"):
        return pd.read_csv(csv_file, delimiter=delimiter)
