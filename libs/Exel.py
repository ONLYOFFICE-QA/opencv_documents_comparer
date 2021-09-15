import csv
import io

from rich import print
from rich.progress import track

from libs.helper import Helper
from var import *


class Exel(Helper):

    def __init__(self, list_of_files):
        self.create_project_dirs()
        # self.copy_for_test(list_of_files)
        self.coordinate = []
        self.errors = []
        self.run_compare_exel(list_of_files)

    @staticmethod
    def get_exel_metadata(wb):
        statistics_exel = {
            'num_of_sheets': f'{wb.Sheets.Count}',
        }
        for sh in wb.Sheets:
            ws = wb.Worksheets(sh.Name)
            used = ws.UsedRange
            nrows = used.Row + used.Rows.Count - 1
            ncols = used.Column + used.Columns.Count - 1
            statistics_exel[f'{sh.Name}_nrows'] = nrows
            statistics_exel[f'{sh.Name}_ncols'] = ncols
        return statistics_exel
        pass

    @staticmethod
    def opener_exel(path_for_open, file_name):
        from win32com.client import Dispatch

        xl = Dispatch("Excel.Application")
        xl.Visible = False  # otherwise excel is hidden

        wb = xl.Workbooks.Open(f'{path_for_open}{file_name}')
        statistics_exel = Exel.get_exel_metadata(wb)
        print("count of sheets:", wb.Sheets.Count)

        wb.Close()
        xl.Quit()
        return statistics_exel

    def run_compare_exel(self, list_of_files):
        with io.open('./report.csv', 'w', encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerow(['File_name', 'statistic'])
            for file_name in track(list_of_files, description='Comparing Exel Metadata...'):
                file_name_from = file_name.replace(f'.{extension_to}', f'.{extension_from}')
                name_to_for_test = self.preparing_file_names(file_name)
                name_from_for_test = self.preparing_file_names(file_name_from)
                if extension_to == file_name.split('.')[-1]:
                    print(f'[bold green]In test[/bold green] {file_name}')
                    print(f'[bold green]In test[/bold green] {file_name_from}')
                    self.copy(f'{custom_doc_to}{file_name}',
                              f'{path_to_temp_in_test}{name_to_for_test}')
                    self.copy(f'{custom_doc_from}{file_name_from}',
                              f'{path_to_temp_in_test}{name_from_for_test}')

                    statistics_exel_after = self.opener_exel(path_to_temp_in_test, name_from_for_test)
                    statistics_exel_before = self.opener_exel(path_to_temp_in_test, name_to_for_test)

                    if statistics_exel_after == {} or statistics_exel_before == {}:
                        print('[bold red]NOT TESTED, Statistics empty!!![/bold red]')
                        self.copy_to_not_tested(file_name,
                                                file_name_from)

                    else:
                        modified = self.dict_compare(statistics_exel_before, statistics_exel_after)
                        if modified != {}:
                            print(modified)
                            self.copy_to_errors(file_name,
                                                file_name_from)
                            modified_keys = [file_name, modified]
                            writer.writerow(modified_keys)

            self.delete(path_to_temp_in_test)
