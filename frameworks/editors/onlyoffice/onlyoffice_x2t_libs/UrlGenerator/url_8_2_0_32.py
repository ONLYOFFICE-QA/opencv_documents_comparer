# -*- coding: utf-8 -*-
from host_tools import HostInfo
from ...handlers.VersionHandler import VersionHandler


class Url82032:

    def __init__(self, version: VersionHandler, host: str):
        self.version = version
        self.host = host
        self.url_version = f"v{self.version.major}.{self.version.minor}"
        self.url_build = self._url_build()
        self.branch = self.branch()

    @property
    def url(self) -> str:
        return f"{self.host}/archive/{self.branch}/{self.url_version}/{self.version.build}/{self.core_name}"

    @property
    def core_name(self) -> str:
        return f"core-{self.os}-{self.arch}.7z"

    def branch(self) -> str:
        if "99.99.99" in self.version.version:
            return 'develop'
        elif self.version.minor != '0':
            return "hotfix"
        return "release"

    def _url_build(self) -> str:
        if HostInfo().os.lower() == 'windows':
            return self.version.version
        return f"{self.version.major}.{self.version.minor}-{self.version.build}"

    @property
    def arch(self) -> str:
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
        os_info = HostInfo().os
        if os_info.lower() in ['linux', 'mac', 'windows']:
            if os_info.lower() == 'windows':
                return 'win'
            return os_info.lower()
        raise ValueError("Unsupported operating system.")
