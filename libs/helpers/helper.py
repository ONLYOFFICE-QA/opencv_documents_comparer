# -*- coding: utf-8 -*-
import os
import json
import random
import shutil
import subprocess as sb
import sys
import codecs
import psutil
import pyautogui as pg
from loguru import logger
from rich import print

from config import *


class Helper:
    def __init__(self, source_extension, converted_extension):
        self.converted_extension = converted_extension
        self.source_extension = source_extension
        self.source_file = None
        self.converted_file = None
        self.tmp_name_converted_file = None
        self.tmp_name_source_file = None
        self.tmp_name = None

        self.ms_office = ms_office
        self.libre_office = libre_office
        self.word = 'WINWORD.EXE'
        self.power_point = 'POWERPNT.EXE'
        self.exel = 'EXCEL.EXE'

        self.project_folder = os.getcwd()
        self.data = self.project_folder + '/data/'
        self.exception_files = self.read_json(f"{self.project_folder}/libs/helpers/exception_file.json")

        self.source_doc_folder = f'{source_doc_folder}{self.source_extension}/'
        self.converted_doc_folder = f'{converted_doc_folder}{version}_{self.source_extension}' \
                                    f'_{self.converted_extension}/'

        self.result_folder = f'{self.data}{version}_{self.source_extension}_{self.converted_extension}/'
        self.differences_statistic = f'{self.result_folder}differences_statistic/'
        self.differences_compare_image = f'{self.result_folder}differences_compare_image/'
        self.passed = f'{self.result_folder}passed/'

        self.untested_folder = f'{self.result_folder}failed_to_open_converted_file/'
        self.failed_source = f'{self.result_folder}failed_to_open_source_file/'
        self.opener_errors = f'{self.result_folder}opener_errors_{converted_extension}_version_{version}/'

        # tmp
        self.tmp_dir = self.data + 'tmp/'
        self.tmp_dir_converted_image = f"{self.tmp_dir}converted_image/"
        self.tmp_dir_source_image = f"{self.tmp_dir}source_image/"
        self.tmp_dir_in_test = f"{self.tmp_dir}in_test/"

        # Create loggers
        logger.remove()
        logger.add(sys.stdout)
        logger.add(f'{self.result_folder}{source_extension}_{converted_extension}_{version}.log',
                   format="{time} {level} {message}",
                   level="DEBUG",
                   rotation='5 MB',
                   compression='zip')

        self.create_project_dirs()
        self.tmp_cleaner()

    @staticmethod
    def read_json(path_to_json):
        with codecs.open(path_to_json, "r", "utf_8_sig") as file:
            json_data = json.load(file)
        return json_data

    @staticmethod
    def click(path):
        try:
            pg.click(path)
        except TypeError:
            pass

    # path insert with file name
    @staticmethod
    def copy(path_from, path_to):
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.copyfile(path_from, path_to)

    @staticmethod
    def random_name(path_for_check, extension):
        while True:
            name = f'{random.randint(5000, 50000000)}.{extension}'
            if not os.path.exists(f'{path_for_check}{name}'):
                return name

    def preparing_files_for_test(self):
        self.source_file = self.converted_file.replace(f'.{self.converted_extension}', f'.{self.source_extension}')

        self.tmp_name_converted_file = self.random_name(self.tmp_dir_in_test, self.converted_extension)
        self.tmp_name_source_file = self.tmp_name_converted_file.replace(f'.{self.converted_extension}',
                                                                         f'.{self.source_extension}')

        if self.source_extension == 'odp':
            self.tmp_name = self.random_name(self.tmp_dir_in_test, self.converted_extension)
            self.copy(f'{self.converted_doc_folder}{self.converted_file}', f'{self.tmp_dir_in_test}{self.tmp_name}')
        else:
            self.tmp_name = self.random_name(self.tmp_dir_in_test, self.source_extension)
            self.copy(f'{self.source_doc_folder}{self.source_file}', f'{self.tmp_dir_in_test}{self.tmp_name}')

        self.copy(f'{self.converted_doc_folder}{self.converted_file}',
                  f'{self.tmp_dir_in_test}{self.tmp_name_converted_file}')
        self.copy(f'{self.source_doc_folder}{self.source_file}',
                  f'{self.tmp_dir_in_test}{self.tmp_name_source_file}')

    @staticmethod
    def move(path_from, path_to):
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.move(path_from, path_to)

    @staticmethod
    def create_dir(path_to_dir):
        if not os.path.exists(path_to_dir):
            os.makedirs(path_to_dir)

    @staticmethod
    def delete(what_delete):
        sb.call(f'powershell.exe rm {what_delete} -Force -Recurse', shell=True)

    @staticmethod
    def dict_compare(source_statistics, converted_statistics):
        d1_keys = set(source_statistics.keys())
        d2_keys = set(converted_statistics.keys())
        shared_keys = d1_keys.intersection(d2_keys)
        modified = {o: (f'Source {source_statistics[o]}', f'Converted {converted_statistics[o]}')
                    for o in shared_keys if source_statistics[o] != converted_statistics[o]}
        return modified

    def run(self, path, file_name, office):
        sb.Popen([f"{self.ms_office}{office}", '-t', f"{path}{file_name}"])

    def run_libre_with_cmd(self, path, file_name):
        sb.Popen([f"{self.libre_office}\simpress.exe", '-o', f"{path}{file_name}"])

    @staticmethod
    def terminate_process():
        terminate_process_list = ['WINWORD.EXE', 'POWERPNT.EXE', 'EXCEL.EXE', 'soffice.bin']
        for process in psutil.process_iter():
            for terminate_process in terminate_process_list:
                if terminate_process in process.name():
                    process.terminate()

    def tmp_cleaner(self):
        self.terminate_process()
        self.delete(f'{self.tmp_dir_in_test}*')
        self.delete(f'{self.tmp_dir_converted_image}*')
        self.delete(f'{self.tmp_dir_source_image}*')

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

    def copy_to_folder(self, path_to_folder):
        if self.converted_file is not None and self.source_file is not None:
            self.create_dir(path_to_folder)
            self.copy(f'{self.converted_doc_folder}{self.converted_file}',
                      f'{path_to_folder}{self.converted_file}')

            self.copy(f'{self.source_doc_folder}{self.source_file}',
                      f'{path_to_folder}{self.source_file}')
        else:
            logger.debug(f'Filename is not found')

    def create_project_dirs(self):
        self.create_dir(self.result_folder)
        self.create_dir(self.tmp_dir_source_image)
        self.create_dir(self.tmp_dir_converted_image)
        self.create_dir(self.tmp_dir_in_test)

