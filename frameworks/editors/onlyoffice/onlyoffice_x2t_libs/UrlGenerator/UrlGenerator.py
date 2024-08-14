# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath
from host_tools import  File
from packaging import version

from ...handlers.VersionHandler import VersionHandler
from .url_8_2_0_32 import Url82032
from .url_7_0_0 import Url700


class UrlGenerator:
    """
    A class to generate URLs for accessing specific build archives based on the provided version.
    The class determines which URL generation strategy to use based on the version number.
    """
    v_8_2_0_32 = version.parse("8.2.0.32")

    def __init__(self, version: str):
        """
        :param version: The version string to generate the URL for.
        """
        self.config = File.read_json(join(dirname(realpath(__file__)), 'url_config.json'))
        self.host = self.config['host']
        self.version = VersionHandler(version)
        self.generator = self.get_generator()
        self.url_version = self.generator.url_version
        self.url_build = self.generator.url_build
        self.branch = self.generator.branch
        self.url = self.generator.url

    def get_generator(self):
        """
        Determines the appropriate URL generator class to use based on the version.

        If the version is 8.2.0.32 or higher, it uses the Url82032 class.
        Otherwise, it falls back to the Url700 class.

        :return: An instance of either Url82032 or Url700 based on the version.
        """
        current_version = version.parse(self.version.version)

        if current_version >= self.v_8_2_0_32:
            return Url82032(version=self.version, host=self.host)
        return Url700(version=self.version, host=self.host)
