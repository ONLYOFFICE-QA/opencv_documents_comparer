# -*- coding: utf-8 -*-
import os
import platform

from rich import print

from framework.StaticData import StaticData
from framework.singleton import singleton


@singleton
class HostActions:
    def __init__(self):
        self.os = ''
        self.arch = ''
        self.x2t = ''
        self.x2ttester = ''
        self.standardtester = ''
        self.host_validations()

    def mac_configuration(self):
        os.environ["DYLD_LIBRARY_PATH"] = StaticData.core_dir()
        self.os = 'mac'
        self.arch = platform.processor()
        self.x2t = 'x2t'
        self.x2ttester = 'x2ttester'
        self.standardtester = 'standardtester'

    def windows_configuration(self):
        self.os = 'windows'
        self.arch = platform.machine().lower().replace("amd", "x")
        self.x2t = 'x2t.exe'
        self.x2ttester = 'x2ttester.exe'
        self.standardtester = 'standardtester.exe'

    def linux_configuration(self):
        self.os = 'linux'
        self.arch = platform.machine().lower().replace("x86_", "x")
        self.x2t = 'x2t'
        self.x2ttester = 'x2ttester'
        self.standardtester = 'standardtester'

    def host_validations(self):
        match platform.system().lower():
            case 'linux':
                self.linux_configuration()
            case 'darwin':
                self.mac_configuration()
            case 'windows':
                self.windows_configuration()
            case _:
                print("[bold red]Error: definition os")
