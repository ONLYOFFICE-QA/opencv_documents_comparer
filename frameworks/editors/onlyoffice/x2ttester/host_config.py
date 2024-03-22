# -*- coding: utf-8 -*-
from os import environ

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from host_tools import HostInfo


@singleton
class HostConfig:
    def __init__(self, dyld_library_path: str = None):
        self.host_info = HostInfo()
        self.dyld_library_path = dyld_library_path or StaticData.core_dir()
        self.arch = self.host_info.arch
        self.os = self.host_info.os
        self.x2t = self.executable_name('x2t')
        self.x2ttester = self.executable_name('x2ttester')
        self.standardtester = self.executable_name('standardtester')
        self.specific_configuration_host()

    def executable_name(self, name):
        if self.os != 'windows':
            return name.replace('.exe', '') if name.endswith('.exe') else name
        return f"{name}.exe"

    def specific_configuration_host(self):
        if self.os == 'mac':
            environ["DYLD_LIBRARY_PATH"] = self.dyld_library_path
