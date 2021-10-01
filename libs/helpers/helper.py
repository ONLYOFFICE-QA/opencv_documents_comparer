import json
import random
import shutil
import subprocess as sb

from rich import print

from variables import *


class Helper:
    # path insert with file name
    @staticmethod
    def copy(path_from, path_to):
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.copyfile(path_from, path_to)

    def preparing_files_for_test(self, converted_file_name, extension_converted, extension_source):
        source_file = converted_file_name.replace(f'.{extension_converted}', f'.{extension_source}')
        tmp_name_converted_file = f'{random.randint(5000, 50000)}.{extension_converted}'
        tmp_name_source_file = tmp_name_converted_file.replace(f'.{extension_converted}',
                                                               f'.{extension_source}')
        tmp_name = f'{random.randint(5000, 50000)}.{extension_source}'
        self.copy(f'{source_doc_folder}{source_file}',
                  f'{tmp_dir_in_test}{tmp_name}')
        self.copy(f'{converted_doc_folder}{converted_file_name}',
                  f'{tmp_dir_in_test}{tmp_name_converted_file}')
        self.copy(f'{source_doc_folder}{source_file}',
                  f'{tmp_dir_in_test}{tmp_name_source_file}')
        return source_file, tmp_name_converted_file, tmp_name_source_file, tmp_name

    @staticmethod
    def move(path_from, path_to):
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.move(path_from, path_to)

    @staticmethod
    def create_dir(path_to_dir):
        if not os.path.exists(path_to_dir):
            os.mkdir(path_to_dir)

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
        sb.Popen([f"{ms_office}{office}", '-t', f"{path}{file_name}"])

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

    def copy_to_folder(self, converted_file, source_file, path_to_folder):
        self.copy(f'{converted_doc_folder}{converted_file}',
                  f'{path_to_folder}{converted_file}')

        self.copy(f'{source_doc_folder}{source_file}',
                  f'{path_to_folder}{source_file}')

    def create_project_dirs(self):
        print('\n')
        self.create_dir(data)
        self.create_dir(result_folder)
        self.create_dir(tmp_dir)
        self.create_dir(passed)
        self.create_dir(tmp_dir_source_image)
        self.create_dir(tmp_dir_converted_image)
        self.create_dir(tmp_dir_in_test)
        self.create_dir(differences_statistic)
        self.create_dir(untested_folder)
        self.create_dir(differences_compare_image)
        self.create_dir(failed_source)
