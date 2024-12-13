# -*- coding: utf-8 -*-
from dataclasses import dataclass
from os import getcwd
from os.path import join
from rich import print

import config


@dataclass(frozen=True)
class StaticData:
    project_dir: str = getcwd()
    logs_dir: str = join(project_dir, 'logs')

    # tmp
    tmp_dir: str = join(project_dir, 'tmp')
    tmp_dir_converted_img: str = join(tmp_dir, 'converted_image')
    tmp_dir_source_img: str = join(tmp_dir, 'source_image')
    tmp_dir_in_test: str = join(tmp_dir, 'in_test')

    word: str = 'WINWORD.EXE'
    powerpoint: str = 'POWERPNT.EXE'
    excel: str = 'EXCEL.EXE'
    libre: str = 'soffice.bin'

    ignore_files = ['.DS_Store']
    ignore_dirs = ['.git', '.idea']

    terminate_process = [word, powerpoint, excel, libre]

    @classmethod
    def documents_dir(cls):
        if not config.source_docs:
            raise print(f"[red]|ERROR| Specify the 'source_docs folder' with the documents to be converted")
        return config.source_docs

    @classmethod
    def core_dir(cls):
        return join(cls.project_dir, 'core')

    @classmethod
    def fonts_dir(cls):
        return join(cls.project_dir, 'assets', 'fonts')

    @classmethod
    def reports_dir(cls):
        return join(cls.project_dir, 'reports')

    @classmethod
    def result_dir(cls):
        return config.converted_docs or join(cls.project_dir, 'result')
