# -*- coding: utf-8 -*-
import re
from os import getcwd
from os.path import join

import settings as st
from framework.FileUtils import FileUtils


class ProjectConfig:
    DOC_ACTIONS = None
    CSTM_REPORT_DIR = None
    PROJECT_DIR = getcwd()
    RESULTS = join(PROJECT_DIR, 'results')
    LOGS_FOLDER = join(PROJECT_DIR, 'logs')
    ASSETS_DIR = join(PROJECT_DIR, 'assets')

    # tmp
    TMP_DIR = join(PROJECT_DIR, 'tmp')
    TMP_DIR_CONVERTED_IMG = join(TMP_DIR, 'converted_image')
    TMP_DIR_SOURCE_IMG = join(TMP_DIR, 'source_image')
    TMP_DIR_IN_TEST = join(TMP_DIR, 'in_test')

    WORD = 'WINWORD.EXE'
    POWERPOINT = 'POWERPNT.EXE'
    EXCEL = 'EXCEL.EXE'
    LIBRE = 'soffice.bin'

    TERMINATE_PROCESS_LIST = [WORD, POWERPOINT, EXCEL, LIBRE]

    EXCEPTION_FILES = FileUtils.read_json(f"{PROJECT_DIR}/data/exception_file.json")

    @staticmethod
    def documents_dir():
        return st.source_docs if st.source_docs else join(ProjectConfig.PROJECT_DIR, 'documents')

    @staticmethod
    def core_dir():
        return join(ProjectConfig.PROJECT_DIR, 'core')

    @staticmethod
    def fonts_dir():
        return join(ProjectConfig.PROJECT_DIR, 'assets', 'fonts')

    @staticmethod
    def reports_dir():
        return join(ProjectConfig.PROJECT_DIR, 'reports', re.sub(r'(\d+).(\d+).(\d+).(\d+)', r'\1.\2.\3', st.version))

    @staticmethod
    def all_fonts_js():
        return join(ProjectConfig.core_dir(), 'fonts', 'AllFonts.js')

    @staticmethod
    def tmp_result_dir():
        return join(ProjectConfig.TMP_DIR, 'cnv')

    @staticmethod
    def core_archive():
        return join(ProjectConfig.TMP_DIR, 'core.7z')

    @staticmethod
    def tools_dir():
        return join(ProjectConfig.PROJECT_DIR, 'tools')

    @staticmethod
    def result_dir():
        return st.converted_docs if st.converted_docs else join(ProjectConfig.result_dir(), 'conversion_result')
