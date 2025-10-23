# -*- coding: utf-8 -*-
from host_tools import HostInfo
from ...handlers.VersionHandler import VersionHandler


class Url700:
    """
    Class for generating download URLs based on the version, operating system, and architecture.
    Attributes:
        host (str): Base URL host.
        url_version (str): Version part of the URL.
        branch (str): Branch of the repository.
    """
    def __init__(self,version: VersionHandler, host: str):
        """
        Initializes the UrlGenerator object.
        :param version: Version of the core.
        """
        self.version = version
        self.host = host
        self.url_version = f"v{self.version.major}.{self.version.minor}"
        self.url_build = self._url_build()

    @property
    def url(self) -> str:
        """
        Generates the download URL.
        :return: Download URL for the core.
        """
        if self.branch == 'develop':
            return f"{self.host}/{self.os}/core/{self.branch}/{self.url_build}/{self.arch}/core.7z"
        return f"{self.host}/{self.os}/core/{self.branch}/{self.url_version}/{self.url_build}/{self.arch}/core.7z"

    @property
    def branch(self) -> str:
        """
        Determines the branch based on the version.
        :return: Branch name.
        """
        return self.version.get_branch()

    @property
    def arch(self) -> str:
        """
        Returns the architecture string based on the operating system.
        :return: Architecture string.
        """
        arch = HostInfo().arch
        if self.os == 'linux':
            return arch.replace("x86_", "x")
        elif self.os == 'mac':
            return arch.replace("64", '')
        elif self.os == 'windows':
            return arch.replace("amd", "x")

    @property
    def os(self) -> str:
        """
        Returns the operating system string.
        :return: Operating system string.
        """
        os_info = HostInfo().os
        if os_info.lower() in ['linux', 'mac', 'windows']:
            return os_info.lower()
        raise ValueError("Unsupported operating system.")

    def _url_build(self) -> str:
        """
        Generates the URL build string.
        :return: URL build string.
        """
        if self.os == 'windows':
            return self.version.version
        return f"{self.version.major}.{self.version.minor}-{self.version.build}"
