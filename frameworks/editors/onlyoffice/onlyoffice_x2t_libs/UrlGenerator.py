# -*- coding: utf-8 -*-
from host_tools import HostInfo
from ..handlers.VersionHandler import VersionHandler


class UrlGenerator(VersionHandler):
    def __init__(self, version):
        super().__init__(version)
        self.host = 'https://s3.eu-west-1.amazonaws.com/repo-doc-onlyoffice-com'
        self.url_version = f"v{self.major_version}.{self.minor_version}"
        self.url_build = self.url_build()
        self.branch = self.branch()

    @property
    def url(self):
        if self.branch == 'develop':
            return f"{self.host}/{self.os}/core/{self.branch}/{self.url_build}/{self.arch}/core.7z"
        return f"{self.host}/{self.os}/core/{self.branch}/{self.url_version}/{self.url_build}/{self.arch}/core.7z"

    def branch(self) -> str:
        if "99.99.99" in self.version:
            return 'develop'
        elif self.minor_version != '0':
            return "hotfix"
        return "release"

    def url_build(self) -> str:
        if self.os == 'windows':
            return self.version
        return f"{self.major_version}.{self.minor_version}-{self.build}"

    @property
    def arch(self) -> str:
        arch = HostInfo().arch
        if self.os == 'linux':
            return arch.replace("x86_", "x")
        elif self.os == 'mac':
            return arch.replace("64", '')
        elif self.os == 'windows':
            return arch.replace("amd", "x")

    @property
    def os(self) -> str:
        os_info = HostInfo().os
        if os_info.lower() in ['linux', 'mac', 'windows']:
            return os_info.lower()
        raise ValueError("Unsupported operating system.")
