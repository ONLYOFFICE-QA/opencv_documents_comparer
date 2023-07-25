# -*- coding: utf-8 -*-
from os import environ

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from frameworks.host_control import HostInfo


@singleton
class HostConfig:
    def __init__(self, dyld_library_path: str = StaticData.core_dir()):
        self.dyld_library_path = dyld_library_path
        self.arch = HostInfo().arch
        self.x2t = self.executable_name('x2t')
        self.x2ttester = self.executable_name('x2ttester')
        self.standardtester = self.executable_name('standardtester')
        self.specific_configuration_host()

    @staticmethod
    def executable_name(name):
        if HostInfo().os != 'windows':
            return name.replace('.exe', '') if name.endswith('.exe') else name
        return f"{name}.exe"

    def specific_configuration_host(self):
        match HostInfo().os:
            case 'mac':
                environ["DYLD_LIBRARY_PATH"] = self.dyld_library_path
