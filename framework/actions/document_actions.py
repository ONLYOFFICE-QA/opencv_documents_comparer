# -*- coding: utf-8 -*-
import sys
from os import walk, listdir
from os.path import join, basename

import psutil
import pyautogui as pg
import pyperclip as pc
from loguru import logger

import settings as config
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.telegram import Telegram
from settings import version, converted_docs, source_docs


class DocActions:
    def __init__(self, source_extension, converted_extension):
        self.source_extension, self.converted_extension = source_extension, converted_extension
        # file_names
        self.source_file, self.converted_file = '', ''
        self.tmp_converted_file, self.tmp_source_file, self.tmp_file_for_get_statistic = '', '', ''
        # paths to document folders
        self.source_doc_folder: str = join(source_docs, source_extension)
        self.converted_doc_folder: str = join(converted_docs, f"{version}_{source_extension}_{converted_extension}")
        # results dirs
        self.result_folder: str = join(ProjectConfig.RESULTS, f"{version}_{source_extension}_{converted_extension}")
        self.differences_statistic: str = join(self.result_folder, 'diff_statistic')
        self.untested_folder: str = join(self.result_folder, 'failed_to_open_converted_file')
        self.failed_source: str = join(self.result_folder, 'failed_to_open_source_file')
        self.opener_errors: str = join(self.result_folder, f"opener_errors_{converted_extension}_version_{version}")
        self.too_long_to_open_files: str = join(self.opener_errors, 'too_long_to_open_files')
        self.create_logger()

    def create_logger(self):
        logger.remove()
        logger.add(sys.stdout)
        logger.add(join(ProjectConfig.LOGS_DIR, f'{self.source_extension}_{self.converted_extension}_{version}.log'),
                   format="{time} {level} {message}",
                   level="DEBUG",
                   rotation='5 MB',
                   compression='zip')

    @staticmethod
    def copy_for_test(path_to_files):
        tmp_file_path = FileUtils.random_name(ProjectConfig.TMP_DIR_IN_TEST, path_to_files.split(".")[-1])
        FileUtils.copy(path_to_files, tmp_file_path)
        return tmp_file_path

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
                        logger.debug(f'Exception when terminate_process: {e}\nFile name: {self.converted_file}')

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
        try:
            pg.click(join(ProjectConfig.PROJECT_DIR, 'data', 'image_templates', path_to_image))
            return True
        except TypeError:
            return False

    @staticmethod
    def last_modified_report():
        return FileUtils.last_modified_file(ProjectConfig.conversion_report_dir())

    @staticmethod
    def copy_result_x2ttester(path_to, output_format, delete=False):
        FileUtils.create_dir(path_to)
        if output_format in ["png", "jpg"]:
            for root, dirs, files in walk(ProjectConfig.tmp_result_dir()):
                for dir_name in dirs:
                    if output_format and dir_name.lower().endswith(f".{output_format.lower()}"):
                        FileUtils.copy(join(root, dir_name), join(path_to, dir_name), silence=True)
        else:
            for file_path in FileUtils.get_files_by_extensions(ProjectConfig.tmp_result_dir(), f".{output_format}"):
                FileUtils.copy(file_path, join(path_to, basename(file_path)), silence=True)
        FileUtils.delete(ProjectConfig.tmp_result_dir()) if delete else None

    def generate_file_array(self, ls=False, df=False, cl=False):
        if ls:
            files_array = config.files_array
        elif cl:
            files_array = pc.paste().split("\n")
        elif df:
            files_array = (listdir(self.differences_statistic))
        else:
            files_array = listdir(self.converted_doc_folder)
        return files_array
