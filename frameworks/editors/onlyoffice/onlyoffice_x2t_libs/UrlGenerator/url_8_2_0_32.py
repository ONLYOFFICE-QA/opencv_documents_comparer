# -*- coding: utf-8 -*-
from host_tools import HostInfo
from ...handlers.VersionHandler import VersionHandler


class Url82032:
    """
    A class to generate a structured URL for accessing specific build archives
    based on version information and the host system's architecture and operating system.
    """

    def __init__(self, version: VersionHandler, host: str):
        """
        :param version: VersionHandler instance containing version details.
        :param host: The base URL or hostname for accessing the archive.
        """
        self.version = version
        self.host = host
        self.url_version = f"v{self.version.major}.{self.version.minor}"
        self.url_build = self._url_build()
        self.branch = self.branch()

    @property
    def url(self) -> str:
        """
        Constructs and returns the full URL for accessing the build archive.
        :return: The complete URL as a string.
        """
        return f"{self.host}/archive/{self.branch}/{self.url_version}/{self.version.build}/{self.core_name}"

    @property
    def core_name(self) -> str:
        """
        Generates the core archive filename based on the OS and architecture.

        :return: The core archive filename as a string.
        """
        return f"core-{self.os}-{self.arch}.7z"

    def branch(self) -> str:
        """
        Determines the appropriate branch based on the version information.

        :return: The branch name ('develop', 'hotfix', or 'release') as a string.
        """
        if "99.99.99" in self.version.version:
            return 'develop'
        elif self.version.minor != '0':
            return "hotfix"
        return "release"

    def _url_build(self) -> str:
        """
        Constructs the URL build segment depending on the host operating system.

        :return: The formatted build string.
        """
        if HostInfo().os.lower() == 'windows':
            return self.version.version
        return f"{self.version.major}.{self.version.minor}-{self.version.build}"

    @property
    def arch(self) -> str:
        """
        Retrieves and processes the architecture information from the host.

        :return: The processed architecture string.
        """
        arch = HostInfo().arch
        os = HostInfo().os.lower()

        if os == 'linux':
            return arch.replace("x86_", "")
        elif os == 'mac':
            return arch.replace("64", '')
        elif os == 'windows':
            return arch.replace("amd", "")

    @property
    def os(self) -> str:
        """
        Retrieves and normalizes the operating system information from the host.

        :return: The normalized OS name as a string.
        :raises ValueError: If the operating system is not supported.
        """
        os_info = HostInfo().os
        if os_info.lower() in ['linux', 'mac', 'windows']:
            if os_info.lower() == 'windows':
                return 'win'
            return os_info.lower()
        raise ValueError("Unsupported operating system.")
