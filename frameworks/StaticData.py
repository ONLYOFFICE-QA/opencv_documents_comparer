# -*- coding: utf-8 -*-
import functools
from os import getcwd
from os.path import join

import settings
from frameworks.host_control.FileUtils import FileUtils


class StaticData:
    DOC_ACTIONS = None
    PROJECT_DIR = getcwd()
    RESULTS = join(PROJECT_DIR, 'results')
    LOGS_DIR = join(PROJECT_DIR, 'logs')

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

    @classmethod
    def documents_dir(cls):
        return settings.source_docs if settings.source_docs else join(cls.PROJECT_DIR, 'documents')

    @classmethod
    def core_dir(cls):
        return join(cls.PROJECT_DIR, 'core')

    @classmethod
    def fonts_dir(cls):
        return join(cls.PROJECT_DIR, 'assets', 'fonts')

    @classmethod
    def reports_dir(cls):
        return join(cls.PROJECT_DIR, 'reports')

    @classmethod
    def result_dir(cls):
        return settings.converted_docs if settings.converted_docs else join(cls.PROJECT_DIR, 'result')
