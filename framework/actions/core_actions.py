# -*- coding: utf-8 -*-
import subprocess as sb
from os import listdir, scandir, chdir
from os.path import join, isdir, isfile, basename
from re import sub

from requests import head
from rich import print
from rich.progress import track

import settings
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions
from framework.actions.xml_actions import XmlActions


class CoreActions:
    def __init__(self, version=''):
        self.xml = XmlActions()
        self.host = HostActions()
        self.os = self.host.os
        self.full_version = self.check_config_version(version if version else settings.version)
        self.branch = self.generate_branch()
        self.major_version = self.generate_major_version(self.full_version)
        self.build = self.generate_build(self.full_version)
        self.url_build = self.generate_build_for_url()
        self.url = self.generate_url()
        self.branches = ["release", "hotfix", "develop"]

    @staticmethod
    def check_config_version(full_version):
        version = full_version if full_version else input("Please enter version core: ")
        if len([i for i in version.split('.') if i]) == 4:
            return version
        raise print("|WARNING| The version is entered incorrectly")

    def generate_build_num(self):
        return self.full_version.split('.')[-1]

    def generate_branch(self):
        if sub(r'(\d+).(\d+).(\d+).(\d+)', r'\3', self.full_version) != '0':
            return "hotfix"
        return 'develop' if "99.99.99" in self.full_version else "release"

    @staticmethod
    def generate_build(full_version):
        if len([i for i in full_version.split('.') if i]) == 4:
            return int(sub(r'(\d+).(\d+).(\d+).(\d+)', r'\4', full_version))
        raise print("[bold red]|WARNING| The version is entered incorrectly")

    def generate_build_for_url(self):
        if self.os == 'windows':
            return self.full_version
        return f"{self.major_version}-{self.build}"

    @staticmethod
    def generate_major_version(full_version):
        if len([i for i in full_version.split('.') if i]) == 4:
            return sub(r'(\d+).(\d+).(\d+).(\d+)', r'\1.\2.\3', full_version)
        raise print("[bold red]|WARNING| The version is entered incorrectly")

    def generate_url(self):
        host = 'https://s3.eu-west-1.amazonaws.com/repo-doc-onlyoffice-com'
        if self.branch == 'develop':
            return f"{host}/{self.os}/core/{self.branch}/{self.url_build}/{self.host.arch}/core.7z"
        return f"{host}/{self.os}/core/{self.branch}/v{self.major_version}/{self.url_build}/{self.host.arch}/core.7z"

    def check_on_server(self):
        self.branches.remove(self.branch) if self.branch in self.branches else ...
        self.branches.insert(0, self.branch)
        for branch in self.branches:
            if self.branch != branch:
                self.branch = branch
                self.url = self.generate_url()
            core_status = head(self.url)
            if core_status.status_code == 200:
                return core_status
        print(f"[bold red]|WARNING| Core not found\nURL:{self.url}\nResponse: {head(self.url).status_code}")
        return False

    def change_core_access(self):
        if self.os == 'windows':
            return
        sb.call(f'chmod +x {ProjectConfig.core_dir()}/*', shell=True)

    @staticmethod
    def write_core_date_on_file(core_data):
        FileUtils.file_writer(join(ProjectConfig.core_dir(), 'core.data'), core_data, mode='w')

    @staticmethod
    def read_core_data(core_path):
        if not isfile(join(core_path, 'core.data')):
            return None
        return FileUtils.file_reader(join(core_path, 'core.data'), mode='r')

    def check_updated_core(self, core_data=None, force=False):
        existing_core_data = self.read_core_data(ProjectConfig.core_dir())
        if core_data and existing_core_data and core_data == existing_core_data and not force:
            print('[red]|INFO| Core Already up-to-date[/]')
            return True

        chdir(ProjectConfig.PROJECT_DIR)
        FileUtils.delete(ProjectConfig.core_dir(), silence=True)
        return False

    def download_core(self):
        print(f"[green]|INFO| Downloading core\nVersion: {self.full_version}\nOS: {self.os}\nURL: {self.url}")
        FileUtils.download_file(self.url, ProjectConfig.TMP_DIR, "core.7z")

    @staticmethod
    def fix_double_folder(core_path=ProjectConfig.core_dir()):
        path = join(core_path, basename(core_path))
        if isdir(path):
            for file in track(listdir(path), description='[green]Fixing the double folder...'):
                FileUtils.move(join(path, file), join(ProjectConfig.core_dir(), file))
            if not any(scandir(path)):
                return FileUtils.delete(path)
            print("[red]|WARNING| Not all objects are moved")

    def getting_core(self, force=False):
        core_status, _ = self.check_on_server(), print('[green]-' * 90)
        if not core_status or self.check_updated_core(core_data=core_status.headers['Last-Modified'], force=force):
            return
        self.download_core()
        FileUtils.unpacking_via_7zip(ProjectConfig.core_archive(), ProjectConfig.core_dir(), delete=True)
        self.fix_double_folder(ProjectConfig.core_dir())
        self.write_core_date_on_file(core_status.headers['Last-Modified'])
        self.change_core_access()
        self.xml.generate_doc_renderer_config(), print('[green]-' * 90)
