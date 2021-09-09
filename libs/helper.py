import json
import shutil
import subprocess as sb

from rich import print
from rich.progress import track

from var import *


class Helper:
    # path insert with file name
    @staticmethod
    def copy(path_from, path_to):
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.copyfile(path_from, path_to)
        else:
            print(f'Folder exists: {path_from}')

    @staticmethod
    def create_dir(path_to_dir):
        if not os.path.exists(path_to_dir):
            os.mkdir(path_to_dir)
        # else:
        #     print('')
        #     # print(f'[bold red]Folder exists: [/bold red]{path_to_dir}')

    # принимает к каталогу с фалвми которые нужно переименовать
    @staticmethod
    def rename_files(dir_for_rename_files):
        for file in track(os.listdir(dir_for_rename_files)):
            # print(os.listdir(dir_for_rename_files))
            extension = file.split('.')[-1]
            file_name = file.replace(f'.{extension}', '')
            file_name = file_name.replace(')', '_')
            file_name = file_name.replace('(', '_')
            file_name = file_name.replace(' ', '_')
            if file != file_name + f'.{extension}':
                # shutil.move(f'{files_for_test_path}{file}', f'{files_for_test_path}{folder_name}{file[-5:]}')
                os.rename(f'{dir_for_rename_files}{file}', f'{dir_for_rename_files}{file_name}.{extension}')

    @staticmethod
    def preparing_file_names(file_name):
        file_name_for_test = file_name.replace(')', '_')
        file_name_for_test = file_name_for_test.replace('(', '_')
        file_name_for_test = file_name_for_test.replace(' ', '_')
        return file_name_for_test

    @staticmethod
    def delete(what_delete):
        sb.call(f'powershell.exe rm {what_delete} -Force -Recurse', shell=True)

    @staticmethod
    def dict_compare(d1, d2):
        d1_keys = set(d1.keys())
        d2_keys = set(d2.keys())
        shared_keys = d1_keys.intersection(d2_keys)
        modified = {o: (f'Before {d1[o]}', f'After {d2[o]}') for o in shared_keys if d1[o] != d2[o]}
        return modified

    @staticmethod
    def run(path, file_name, office):
        sb.Popen([f"{ms_office}{office}", f"{path}{file_name}"])

    @staticmethod
    def create_word_json(word, file_name_for_save, path_for_save):
        statistics_word = {
            'num_of_sheets': f'{word.ComputeStatistics(2)}',
            'number_of_lines': f'{word.ComputeStatistics(1)}',
            'word_count': f'{word.ComputeStatistics(0)}',
            'number_of_characters_without_spaces': f'{word.ComputeStatistics(3)}',
            'number_of_characters_with_spaces': f'{word.ComputeStatistics(5)}',
            'number_of_abzad': f'{word.ComputeStatistics(4)}',
        }
        with open(f'{path_for_save}{file_name_for_save}_doc.json', 'w') as f:
            json.dump(statistics_word, f)
        return statistics_word

    def copy_for_test(self, list_of_files=list_file_names_doc_from_compare, ext_from=extension_from,
                      ext_to=extension_to):
        for file_name in list_of_files:
            self.copy(f'{custom_doc_to}{file_name}',
                      f'{path_to_folder_for_test}{file_name}')
            file_name_from = file_name.replace(f'.{ext_to}', f'.{ext_from}')
            self.copy(f'{custom_doc_from}{file_name_from}',
                      f'{path_to_folder_for_test}{file_name_from}')

    def create_project_dirs(self):
        print('')
        self.create_dir(path_to_data)
        self.create_dir(path_to_tmp)
        self.create_dir(path_to_result)
        self.create_dir(path_to_compare_files)
        self.create_dir(path_to_tmpimg_befor_conversion)
        self.create_dir(path_to_tmpimg_after_conversion)
        self.create_dir(path_to_temp_in_test)
        self.create_dir(path_to_folder_for_test)
        self.create_dir(path_to_errors_file)
        self.create_dir(path_to_not_tested_file)
