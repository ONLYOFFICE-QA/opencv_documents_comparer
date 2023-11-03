# -*- coding: utf-8 -*-
from re import sub
from rich import print


class VersionHandler:
    def __init__(self, version: str):
        self._version_pattern = r'(\d+).(\d+).(\d+).(\d+)'
        self.version = version

    @property
    def version(self):
        return self.__version

    @version.setter
    def version(self, value: str) -> None:
        if len([i for i in value.split('.') if i]) != 4:
            raise print("[red]|WARNING| Version is entered incorrectly.The version must be in the format 'x.x.x.x'")
        self.__version = value

    @property
    def major_version(self) -> str:
        return sub(self._version_pattern, r'\1.\2', self.version)

    @property
    def minor_version(self) -> str:
        return sub(self._version_pattern, r'\3', self.version)

    @property
    def build(self) -> int:
        return int(sub(self._version_pattern, r'\4', self.version))

    @property
    def without_build(self) -> str:
        return f"{self.major_version}.{self.minor_version}"
