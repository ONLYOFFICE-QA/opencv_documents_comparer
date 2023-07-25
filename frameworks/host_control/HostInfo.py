# -*- coding: utf-8 -*-
from platform import system, machine

from ..decorators.decorators import singleton


@singleton
class HostInfo:
    def __init__(self):
        self.os = system().lower()
        self.__arch = machine().lower()

    @property
    def os(self):
        return self.__os

    @property
    def arch(self):
        return self.__arch

    @os.setter
    def os(self, value):
        match value:
            case 'linux':
                self.__os = 'linux'
            case 'darwin':
                self.__os = 'mac'
            case 'windows':
                self.__os = 'windows'
            case definition_os:
                print(f"[bold red]|WARNING| Error defining os: {definition_os}")
