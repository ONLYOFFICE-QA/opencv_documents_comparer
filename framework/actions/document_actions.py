# -*- coding: utf-8 -*-
from loguru import logger
import random
from os import walk, listdir
from os.path import join, exists
import sys
import psutil
import pyperclip as pc
import pyautogui as pg
import configuration as config

from configuration import version, converted_docs, source_docs
from data.project_configurator import ProjectConfig
from framework.telegram import Telegram
from framework.FileUtils import FileUtils


class DocActions:
    def __init__(self, source_extension, converted_extension):
        self.source_extension, self.converted_extension = source_extension, converted_extension
        # file_names
        self.source_file, self.converted_file = '', ''
        self.tmp_converted_file, self.tmp_source_file, self.tmp_file_for_get_statistic = '', '', ''
        # paths to document folders
        self.source_doc_folder: str = f'{source_docs}/{source_extension}/'
        self.converted_doc_folder: str = f'{converted_docs}/{version}_{source_extension}_{converted_extension}/'
        # results dirs
        self.result_folder: str = f'{ProjectConfig.RESULTS}/{version}_{source_extension}_{converted_extension}/'
        self.differences_statistic: str = f'{self.result_folder}/differences_statistic/'
        self.untested_folder: str = f'{self.result_folder}/failed_to_open_converted_file/'
        self.failed_source: str = f'{self.result_folder}/failed_to_open_source_file/'
        self.opener_errors: str = f'{self.result_folder}/opener_errors_{converted_extension}_version_{version}/'
        self.too_long_to_open_files: str = f'{self.opener_errors}/too_long_to_open_files/'
        self.create_logger()
        self.tmp_cleaner()

    def create_logger(self):
        logger.remove()
        logger.add(sys.stdout)
        logger.add(join(ProjectConfig.LOGS_FOLDER, f'{self.source_extension}_{self.converted_extension}_{version}.log'),
                   format="{time} {level} {message}",
                   level="DEBUG",
                   rotation='5 MB',
                   compression='zip')

    @staticmethod
    def random_name(file_extension):
        while True:
            random_file_name = f'{random.randint(5000, 50000000)}.{file_extension}'
            if not exists(join(ProjectConfig.TMP_DIR_IN_TEST, random_file_name)):
                return random_file_name

    @staticmethod
    def copy_for_test(path_to_files):
        tmp_name = DocActions.random_name(path_to_files.split(".")[-1])
        FileUtils.copy(path_to_files, join(ProjectConfig.TMP_DIR_IN_TEST, tmp_name))
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
            for terminate_process in ProjectConfig.TERMINATE_PROCESS_LIST:
                if terminate_process in process.name():
                    try:
                        process.terminate()
                    except Exception as e:
                        message = f'Exception when terminate_process: {e}\nFile name: {self.converted_file}'
                        logger.debug(message)

    def tmp_cleaner(self):
        self.terminate_process()
        FileUtils.delete(f'{ProjectConfig.TMP_DIR_IN_TEST}', all_from_folder=True, silence=True)
        FileUtils.delete(f'{ProjectConfig.TMP_DIR_CONVERTED_IMG}', all_from_folder=True, silence=True)
        FileUtils.delete(f'{ProjectConfig.TMP_DIR_SOURCE_IMG}', all_from_folder=True, silence=True)

    def copy_testing_files_to_folder(self, dir_path):
        if self.converted_file and self.source_file:
            FileUtils.create_dir(dir_path)
            FileUtils.copy(join(self.converted_doc_folder, self.converted_file), join(dir_path, self.converted_file))
            FileUtils.copy(join(self.source_doc_folder, self.source_file), join(dir_path, self.source_file))

    def create_massage_for_tg(self, errors_array, ls):
        massage = f'{self.source_extension}=>{self.converted_extension} ' \
                  f'opening check completed on version: {version}\n' \
                  f'Files with errors when opening:\n`{errors_array}`'
        passed_files = [file for file in config.files_array if file not in errors_array]
        Telegram.send_message(massage) if not ls else print(f'{massage}\n\nPassed files:\n{passed_files}')

    @staticmethod
    def dict_compare(source_statistics, converted_statistics):
        d1_keys, d2_keys = set(source_statistics.keys()), set(converted_statistics.keys())
        modified = {
            o: (f'Source {source_statistics[o]}', f'Converted {converted_statistics[o]}')
            for o in d1_keys.intersection(d2_keys) if source_statistics[o] != converted_statistics[o]
        }
        return modified

    @staticmethod
    def click(path_to_image):
        from data.project_configurator import ProjectConfig
        try:
            pg.click(f'{ProjectConfig.PROJECT_DIR}/data/image_templates/{path_to_image}')
            return True
        except TypeError:
            return False

    @staticmethod
    def last_modified_report():
        return FileUtils.last_modified_file(ProjectConfig.CSTM_REPORT_DIR)

    def prepare_files_for_openers(self, input_format, output_format, x2t_version):
        result_folder = f"{ProjectConfig.result_dir()}/{x2t_version}_{input_format}_{output_format}"
        FileUtils.create_dir(result_folder)
        self.copy_result_x2ttester(result_folder, output_format, delete=False)

    @staticmethod
    def copy_result_x2ttester(path_to, output_format, delete=False):
        for root, dirs, files in walk(f"{ProjectConfig.tmp_result_dir()}"):
            if output_format in ["png", "jpg"]:
                for dirname in dirs:
                    if output_format and dirname.lower().endswith(f".{output_format.lower()}"):
                        FileUtils.copy(join(root, dirname), join(path_to, dirname), silence=True)
            else:
                for filename in files:
                    if output_format and filename.lower().endswith(f".{output_format.lower()}"):
                        FileUtils.copy(join(root, filename), join(path_to, filename), silence=True)
        FileUtils.delete(f"{ProjectConfig.tmp_result_dir()}") if delete else None

    def get_file_array(self, ls=False, df=False, cl=False):
        if ls:
            files_array = config.files_array
        elif cl:
            files_array = pc.paste().split("\n")
        elif df:
            files_array = (listdir(self.differences_statistic))
        else:
            files_array = listdir(self.converted_doc_folder)
        return files_array
