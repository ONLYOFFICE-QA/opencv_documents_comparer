# -*- coding: utf-8 -*-
from libs.helpers.fileutils import FileUtils
from management import *


class StaticData:
    DOC_HELPER = None
    PROJECT_FOLDER = os.getcwd()
    RESULTS = f'{PROJECT_FOLDER}/results/'
    LOGS_FOLDER = f"{PROJECT_FOLDER}/logs/"

    # tmp
    TMP_DIR = f'{PROJECT_FOLDER}/tmp/'
    TMP_DIR_CONVERTED_IMG = f"{TMP_DIR}converted_image/"
    TMP_DIR_SOURCE_IMG = f"{TMP_DIR}source_image/"
    TMP_DIR_IN_TEST = f"{TMP_DIR}in_test/"

    WORD = 'WINWORD.EXE'
    POWERPOINT = 'POWERPNT.EXE'
    EXCEL = 'EXCEL.EXE'
    LIBRE = 'soffice.bin'

    TERMINATE_PROCESS_LIST = [WORD, POWERPOINT, EXCEL, LIBRE]

    EXCEPTION_FILES = FileUtils.read_json(f"{PROJECT_FOLDER}/data/exception_file.json")
