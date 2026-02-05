import pandas as pd
from rich import print
from .x2ttester_conversion import X2tTesterConversion

class ErrorHandler:
    check_exit_codes = ['88', '88.0']

    def __init__(self, conversion: X2tTesterConversion, report: str):
        self.conversion = conversion
        self.report = report

    def handle_errors(self):
        delete_88_files_list = self.get_errors88_files(self.conversion.report.read(self.report))
        print(f"[magenta]|INFO| Files with Exit_code 88 to check: {delete_88_files_list}")

        if delete_88_files_list:
            cheked_files_list = self.check_errors88_files(delete_88_files_list)
            delete_88_files_list = [file for file in delete_88_files_list if file not in cheked_files_list]

        print(f"[bold cyan]|INFO| Files with Exit_code 88 to delete from report: {delete_88_files_list}")
        # Remove files with Exit_code 88 from the main report if they remain after rechecks
        if delete_88_files_list:
            main_report_df = self.conversion.report.read(self.report)
            main_report_df = main_report_df[~main_report_df['Input file'].isin(delete_88_files_list)]
            main_report_df.to_csv(self.report, sep='\t', index=False)

    def check_errors88_files(self, delete_88_files_list):
        report2 = self.conversion.from_files_list(delete_88_files_list)
        for _ in range(5):
            cheked_files_list = self.get_errors88_files(self.conversion.report.read(report2))
            print(f"[yellow]Checked files with Exit_code 88: {cheked_files_list}")
            if not cheked_files_list:
                print(f"[red]Checked files with Exit_code 88 is empty")
                return cheked_files_list
            report2 = self.conversion.from_files_list(cheked_files_list)

        return self.get_errors88_files(self.conversion.report.read(report2))

    def get_errors88_files(self, df: pd.DataFrame):
        _df = df
        _df['Exit code'] = _df['Exit code'].astype(str)
        # Print all unique exit codes
        print("[yellow]Unique exit codes:", _df['Exit code'].unique())
        # Filter for exit codes '88' and '88.0'
        _df = _df[_df['Exit code'].isin(self.check_exit_codes)]
        return _df['Input file'].tolist()
