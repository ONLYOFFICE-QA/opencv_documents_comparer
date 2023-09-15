# -*- coding: utf-8 -*-
from host_control import HostInfo
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

    def branch(self):
        if self.minor_version != '0':
            return "hotfix"
        return 'develop' if "99.99.99" in self.version else "release"

    def url_build(self):
        if self.os == 'windows':
            return self.version
        return f"{self.major_version}.{self.minor_version}-{self.build}"

    @property
    def arch(self):
        match HostInfo().os:
            case 'linux':
                return HostInfo().arch.replace("x86_", "x")
            case 'mac':
                return HostInfo().arch.replace("64", '')
            case 'windows':
                return HostInfo().arch.replace("amd", "x")

    @property
    def os(self):
        match HostInfo().os:
            case 'linux':
                return 'linux'
            case 'mac':
                return 'mac'
            case 'windows':
                return 'windows'
