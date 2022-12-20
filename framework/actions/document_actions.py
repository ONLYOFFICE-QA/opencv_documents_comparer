# -*- coding: utf-8 -*-
from loguru import logger
import random
import os
from os.path import join
import sys
import psutil
import pyperclip as pc
import configuration as config

from configuration import version, converted_doc_folder, source_doc_folder
from data.StaticData import StaticData
from framework.telegram import Telegram
from framework.fileutils import FileUtils


class DocActions:
    def __init__(self, source_extension, converted_extension):
        self.source_extension, self.converted_extension = source_extension, converted_extension
        # file_names
        self.source_file, self.converted_file = '', ''
        self.tmp_converted_file, self.tmp_source_file, self.tmp_file_for_get_statistic = '', '', ''
        # paths to document folders
        self.source_doc_folder: str = f'{source_doc_folder}/{source_extension}/'
        self.converted_doc_folder: str = f'{converted_doc_folder}/{version}_{source_extension}_{converted_extension}/'
        # results dirs
        self.result_folder: str = f'{StaticData.RESULTS}/{version}_{source_extension}_{converted_extension}/'
        self.differences_statistic: str = f'{self.result_folder}/differences_statistic/'
        self.untested_folder: str = f'{self.result_folder}/failed_to_open_converted_file/'
        self.failed_source: str = f'{self.result_folder}/failed_to_open_source_file/'
        self.opener_errors: str = f'{self.result_folder}/opener_errors_{converted_extension}_version_{version}/'
        self.too_long_to_open_files: str = f'{self.opener_errors}/too_long_to_open_files/'
        self.create_logger()
        self.create_tmp_dirs()
        self.tmp_cleaner()

    def create_logger(self):
        logger.remove()
        logger.add(sys.stdout)
        logger.add(join(StaticData.LOGS_FOLDER, f'{self.source_extension}_{self.converted_extension}_{version}.log'),
                   format="{time} {level} {message}",
                   level="DEBUG",
                   rotation='5 MB',
                   compression='zip')

    @staticmethod
    def random_name(file_extension):
        while True:
            random_file_name = f'{random.randint(5000, 50000000)}.{file_extension}'
            if not os.path.exists(join(StaticData.TMP_DIR_IN_TEST, random_file_name)):
                return random_file_name

    @staticmethod
    def copy_for_test(path_to_files):
        tmp_name = DocActions.random_name(path_to_files.split(".")[-1])
        FileUtils.copy(path_to_files, join(StaticData.TMP_DIR_IN_TEST, tmp_name))
        return tmp_name

    def preparing_files_for_opening_test(self):
        self.tmp_converted_file = self.copy_for_test(join(self.converted_doc_folder, self.converted_file))

    def preparing_files_for_compare_test(self):
        self.source_file = self.converted_file.replace(f'.{self.converted_extension}', f'.{self.source_extension}')
        self.tmp_converted_file = self.copy_for_test(join(self.converted_doc_folder, self.converted_file))
        self.tmp_source_file = self.copy_for_test(join(self.source_doc_folder, self.source_file))
        if self.source_extension in ['odp', 'ods', 'odt']:
            self.tmp_file_for_get_statistic = self.copy_for_test(join(self.converted_doc_folder, self.converted_file))
        else:
            self.tmp_file_for_get_statistic = self.copy_for_test(join(self.source_doc_folder, self.source_file))

    def terminate_process(self):
        for process in psutil.process_iter():
            for terminate_process in StaticData.TERMINATE_PROCESS_LIST:
                if terminate_process in process.name():
                    try:
                        process.terminate()
                    except Exception as e:
                        message = f'Exception when terminate_process: {e}\nFile name: {self.converted_file}'
                        logger.debug(message)

    def tmp_cleaner(self):
        self.terminate_process()
        FileUtils.delete(f'{StaticData.TMP_DIR_IN_TEST}', all_from_folder=True)
        FileUtils.delete(f'{StaticData.TMP_DIR_CONVERTED_IMG}', all_from_folder=True)
        FileUtils.delete(f'{StaticData.TMP_DIR_SOURCE_IMG}', all_from_folder=True)

    def copy_testing_files_to_folder(self, dir_path):
        if self.converted_file and self.source_file:
            FileUtils.create_dir(dir_path)
            FileUtils.copy(join(self.converted_doc_folder, self.converted_file), join(dir_path, self.converted_file))
            FileUtils.copy(join(self.source_doc_folder, self.source_file), join(dir_path, self.source_file))
        else:
            logger.debug(f'Filename is not found')

    @staticmethod
    def create_tmp_dirs():
        FileUtils.create_dir(StaticData.TMP_DIR_SOURCE_IMG)
        FileUtils.create_dir(StaticData.TMP_DIR_CONVERTED_IMG)
        FileUtils.create_dir(StaticData.TMP_DIR_IN_TEST)

    def create_massage_for_tg(self, errors_array, ls):
        massage = f'{self.source_extension}=>{self.converted_extension}' \
                  f'opening check completed on version:{version}\n' \
                  f'Files with errors when opening:\n`{errors_array}`'
        passed_files = [file for file in config.list_of_file_names if file not in errors_array]
        Telegram.send_message(massage) if not ls else print(f'{massage}\n\nPassed files:\n{passed_files}')

    @staticmethod
    def dict_compare(source_statistics, converted_statistics):
        d1_keys, d2_keys = set(source_statistics.keys()), set(converted_statistics.keys())
        modified = {
            o: (f'Source {source_statistics[o]}', f'Converted {converted_statistics[o]}')
            for o in d1_keys.intersection(d2_keys) if source_statistics[o] != converted_statistics[o]
        }
        return modified

    def get_file_array(self, ls=False, df=False, cl=False):
        if ls:
            files_array = config.list_of_file_names
        elif cl:
            files_array = pc.paste().split("\n")
        elif df:
            files_array = (os.listdir(self.differences_statistic))
        else:
            files_array = os.listdir(self.converted_doc_folder)
        return files_array
