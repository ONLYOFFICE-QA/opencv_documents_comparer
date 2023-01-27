# -*- coding: utf-8 -*-
import os
import platform

from rich import print

from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
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
        FileUtils.delete(ProjectConfig.TMP_DIR, silence=True)
        self.create_project_dirs()

    @staticmethod
    def create_project_dirs():
        FileUtils.create_dir(ProjectConfig.reports_dir())
        FileUtils.create_dir(ProjectConfig.TMP_DIR_IN_TEST, silence=True)

    def mac_configuration(self):
        os.environ["DYLD_LIBRARY_PATH"] = ProjectConfig.core_dir()
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
