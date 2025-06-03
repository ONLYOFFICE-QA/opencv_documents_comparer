# -*- coding: utf-8 -*-
import csv
import pandas as pd

from os.path import dirname, isfile
from csv import reader

from rich import print

from host_tools import Dir


class Report:
    """
    Utility class for handling reports.
    """

    def __init__(self):
        """
        Initializes the Report object.
        Sets display options for pandas DataFrame.
        """
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option("expand_frame_repr", False)

    def merge(self, reports: list, result_csv_path: str, delimiter='\t') -> str | None:
        """
        Merges multiple CSV reports into a single DataFrame and writes it to a new CSV file.
        :param reports: List of paths to CSV report files.
        :param result_csv_path: Path to the resulting merged CSV file.
        :param delimiter: Delimiter to use for CSV files (default is '\t').
        :return: Path to the resulting merged CSV file if successful, None otherwise.
        """
        if reports:
            df = pd.concat([self.read(csv_, delimiter) for csv_ in reports if isfile(csv_)], ignore_index=True)
            df.to_csv(result_csv_path, index=False, sep=delimiter)
            return result_csv_path

        print('[green]|INFO| No files to merge')

    @staticmethod
    def write(file_path: str, mode: str, message: list, delimiter='\t', encoding='utf-8') -> None:
        """
        Writes a message to a CSV file.
        :param file_path: Path to the CSV file.
        :param mode: File opening mode ('w' for write, 'a' for append).
        :param message: List of data to write to the CSV file.
        :param delimiter: Delimiter to use for the CSV file (default is '\t').
        :param encoding: Encoding to use for the CSV file (default is 'utf-8').
        :return: None
        """
        Dir.create(dirname(file_path), stdout=False)
        with open(file_path, mode, newline='', encoding=encoding) as csv_file:
            writer = csv.writer(csv_file, delimiter=delimiter)
            writer.writerow(message)

    @staticmethod
    def read(csv_file: str, delimiter="\t", **kwargs) -> pd.DataFrame:
        """
        Reads a CSV file into a pandas DataFrame.
        :param csv_file: Path to the CSV file.
        :param delimiter: Delimiter used in the CSV file (default is '\t').
        :return: DataFrame containing the data from the CSV file.
        """
        data = pd.read_csv(csv_file, delimiter=delimiter, **kwargs)
        last_row = data.iloc[-1]

        if last_row.isnull().all() or (last_row.astype(str).str.contains(r"[^\x00-\x7F]", regex=True).any()):
            data = data.iloc[:-1]
            data.to_csv(csv_file)

        return data

    @staticmethod
    def read_via_csv(csv_file: str, delimiter: str = "\t") -> list:
        """
        Reads a CSV file and returns its content as a list of lists.
        :param csv_file: Path to the CSV file.
        :param delimiter: Delimiter used in the CSV file (default is '\t').
        :return: List containing rows from the CSV file.
        """
        with open(csv_file, 'r') as csvfile:
            return [row for row in reader(csvfile, delimiter=delimiter)]

    @staticmethod
    def save_csv(df, csv_path: str, delimiter="\t") -> str:
        """
        Saves a pandas DataFrame to a CSV file.
        :param df: DataFrame to save.
        :param csv_path: Path to save the CSV file.
        :param delimiter: Delimiter to use for the CSV file (default is '\t').
        :return: Path to the saved CSV file.
        """
        df.to_csv(csv_path, index=False, sep=delimiter)
        return csv_path
