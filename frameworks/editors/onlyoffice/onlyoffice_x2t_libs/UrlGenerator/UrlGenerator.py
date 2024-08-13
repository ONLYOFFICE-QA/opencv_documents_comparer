# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath
from host_tools import  File

from frameworks.editors.onlyoffice.handlers.VersionHandler import VersionHandler
from .url_8_2_0_32 import Url82032
from .url_7_0_0 import Url700


class UrlGenerator:

    def __init__(self, version: str):
        self.config = File.read_json(join(dirname(realpath(__file__)), 'url_config.json'))
        self.host = self.config['host']
        self.version = VersionHandler(version)
        self.generator = self.get_generator()
        self.url_version = self.generator.url_version
        self.url_build = self.generator.url_build
        self.branch = self.generator.branch
        self.url = self.generator.url

    def get_generator(self):
        if self.compare_versions(
                version=self.version.version,
                min_version="99.99.99.3950" if "99.99.99" in self.version.version else '8.2.0.32'
        ):
            return Url82032(version=self.version, host=self.host)
        return Url700(version=self.version, host=self.host)

    @staticmethod
    def compare_versions(version: str, min_version: str) -> bool:
        def version_to_tuple(v):
            return tuple(map(int, v.split('.')))

        return version_to_tuple(version) >= version_to_tuple(min_version)
