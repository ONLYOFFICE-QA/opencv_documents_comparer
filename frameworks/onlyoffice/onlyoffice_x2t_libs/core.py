# -*- coding: utf-8 -*-
from os import chdir
from os.path import join, isfile
from .UrlGenerator import UrlGenerator
from rich import print

from frameworks.decorators.decorators import highlighter, timer_decorator
from ..xml.x2t_libs_xml import X2tLibsXML
from ...StaticData import StaticData
from ...host_control import FileUtils, HostInfo


class Core:
    def __init__(self, version: str = None):
        self.os = HostInfo().os
        self.url = UrlGenerator(version).url
        self.version = version
        self.core_dir = StaticData.core_dir()
        self.tmp_dir = StaticData.TMP_DIR
        self.project_dir = StaticData.PROJECT_DIR
        self.data_file = join(self.core_dir, 'core.data')

    def read_core_data(self):
        if not isfile(self.data_file):
            return None
        return FileUtils.file_reader(self.data_file, mode='r')

    def check_updated_core(self, core_data=None):
        existing_core_data = self.read_core_data()
        if core_data and existing_core_data and core_data == existing_core_data:
            print('[red]|INFO| Core Already up-to-date[/]')
            return True
        return False

    def download(self):
        print(f"[green]|INFO| Downloading core\nVersion: {self.version}\nOS: {self.os}\nURL: {self.url}")
        FileUtils.download_file(self.url, self.tmp_dir, "core.7z")

    def delete_core_dir(self):
        chdir(self.project_dir)
        FileUtils.delete(self.core_dir, silence=True)

    @highlighter(color='green')
    def getting(self, force=False):
        self.delete_core_dir() if force else ...
        headers = FileUtils.get_headers(self.url)
        if not headers or self.check_updated_core(core_data=headers['Last-Modified']):
            return
        self.download()
        FileUtils.unpacking_7zip(join(self.tmp_dir, "core.7z"), self.core_dir, delete=True)
        FileUtils.fix_double_folder(self.core_dir)
        FileUtils.change_access(self.core_dir)
        FileUtils.file_writer(self.data_file, headers['Last-Modified'], mode='w')
        X2tLibsXML().create_doc_renderer_config()
