# -*- coding: utf-8 -*-
from host_tools import HostInfo
from ..handlers.VersionHandler import VersionHandler


class UrlGenerator(VersionHandler):
    """
    Class for generating download URLs based on the version, operating system, and architecture.

    Inherits from VersionHandler to handle version-specific logic.

    Attributes:
        host (str): Base URL host.
        url_version (str): Version part of the URL.
        branch (str): Branch of the repository.
    """
    def __init__(self, version: str):
        """
        Initializes the UrlGenerator object.
        :param version: Version of the core.
        """
        super().__init__(version)
        self.host = 'https://s3.eu-west-1.amazonaws.com/repo-doc-onlyoffice-com'
        self.url_version = f"v{self.major_version}.{self.minor_version}"
        self.url_build = self.url_build()
        self.branch = self.branch()

    @property
    def url(self) -> str:
        """
        Generates the download URL.
        :return: Download URL for the core.
        """
        if self.branch == 'develop':
            return f"{self.host}/{self.os}/core/{self.branch}/{self.url_build}/{self.arch}/core.7z"
        return f"{self.host}/{self.os}/core/{self.branch}/{self.url_version}/{self.url_build}/{self.arch}/core.7z"

    def branch(self) -> str:
        """
        Determines the branch based on the version.
        :return: Branch name.
        """
        if "99.99.99" in self.version:
            return 'develop'
        elif self.minor_version != '0':
            return "hotfix"
        return "release"

    def url_build(self) -> str:
        """
        Generates the URL build string.
        :return: URL build string.
        """
        if self.os == 'windows':
            return self.version
        return f"{self.major_version}.{self.minor_version}-{self.build}"

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
