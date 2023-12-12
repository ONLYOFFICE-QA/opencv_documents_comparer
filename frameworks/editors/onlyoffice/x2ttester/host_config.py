# -*- coding: utf-8 -*-
from os import environ

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from host_tools import HostInfo


@singleton
class HostConfig:
    def __init__(self, dyld_library_path: str = StaticData.core_dir()):
        self.dyld_library_path = dyld_library_path
        self.arch = HostInfo().arch
        self.os = HostInfo().os
        self.x2t = self.executable_name('x2t')
        self.x2ttester = self.executable_name('x2ttester')
        self.standardtester = self.executable_name('standardtester')
        self.specific_configuration_host()

    def executable_name(self, name):
        if self.os != 'windows':
            return name.replace('.exe', '') if name.endswith('.exe') else name
        return f"{name}.exe"

    def specific_configuration_host(self):
        match self.os:
            case 'mac':
                environ["DYLD_LIBRARY_PATH"] = self.dyld_library_path
