# -*- coding: utf-8 -*-
from os import environ

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from host_tools import HostInfo


@singleton
class HostConfig:
    """
    Singleton class representing host configuration.

    Provides information about the host environment such as architecture,
    operating system, and paths to various executables.

    Attributes:
        host_info (HostInfo): Object containing information about the host.
        dyld_library_path (str): Path to the dynamic linker library (macOS only).
        arch (str): Architecture of the host system.
        os (str): Operating system of the host system.
        x2t (str): Path to the x2t executable.
        x2ttester (str): Path to the x2ttester executable.
        standardtester (str): Path to the standardtester executable.
    """

    def __init__(self, dyld_library_path: str = None):
        """
        Initializes the HostConfig object.
        :param dyld_library_path: Path to the dynamic linker library (macOS only).
        """
        self.host_info = HostInfo()
        self.dyld_library_path = dyld_library_path or StaticData.core_dir()
        self.arch = self.host_info.arch
        self.os = self.host_info.os
        self.x2t = self.executable_name('x2t')
        self.x2ttester = self.executable_name('x2ttester')
        self.standardtester = self.executable_name('standardtester')
        self.specific_configuration_host()

    def executable_name(self, name):
        """
        Returns the executable name based on the operating system.
        :param name: Original executable name.
        :return: Adjusted executable name based on the current operating system.
        """
        if self.os != 'windows':
            return name.replace('.exe', '') if name.endswith('.exe') else name
        return f"{name}.exe"

    def specific_configuration_host(self):
        """
        Performs specific configuration based on the host operating system.

        Sets environment variables or performs other necessary configurations
        based on the host operating system.
        """
        if self.os == 'mac':
            environ["DYLD_LIBRARY_PATH"] = self.dyld_library_path
