# -*- coding: utf-8 -*-
import re
import subprocess as sb
from os import chdir, listdir, scandir
from os.path import join, isdir, isfile, basename

from requests import head
from rich import print
from rich.progress import track

import settings
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions
from framework.actions.xml_actions import XmlActions


class CoreActions:
    def __init__(self, version=settings.version):
        self.host = HostActions()
        self.xml = XmlActions()
        self.config_version = version if version else input("Please enter version core: ")
        self.branch = self.generate_branch()
        self.os = self.host.os
        self.arch = self.host.arch
        self.build = self.generate_build()
        self.branch_version = self.generate_version()
        self.url = self.generate_url()
        self.branches = ["release", "hotfix", "develop"]

    def generate_build_num(self):
        return self.config_version.split('.')[-1]

    def generate_branch(self):
        return 'develop' if "99.99.99" in self.config_version else "release"

    def generate_build(self):
        if self.os == 'windows':
            return self.config_version
        return re.sub(r'(\d+).(\d+).(\d+).(\d+)', r'\1.\2.\3-\4', self.config_version)

    def generate_version(self):
        return re.sub(r'(\d+).(\d+).(\d+).(\d+)', r'v\1.\2.\3', self.config_version)

    def generate_url(self):
        host = 'https://s3.eu-west-1.amazonaws.com/repo-doc-onlyoffice-com'
        if self.branch == 'develop':
            return f"{host}/{self.os}/core/{self.branch}/{self.build}/{self.arch}/core.7z"
        return f"{host}/{self.os}/core/{self.branch}/{self.branch_version}/{self.build}/{self.arch}/core.7z"

    def check_on_server(self):
        self.branches.remove(self.branch) if self.branch in self.branches else None
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

    def generate_allfonts(self):
        chdir(ProjectConfig.core_dir())
        sb.call(join(ProjectConfig.core_dir(), self.host.standardtester))
        chdir(ProjectConfig.PROJECT_DIR)

    @staticmethod
    def write_core_date_on_file(core_data):
        with open(join(ProjectConfig.core_dir(), 'core.data'), "w") as file:
            file.write(core_data)

    @staticmethod
    def read_core_data(core_path):
        if not isfile(join(core_path, 'core.data')):
            return None
        with open(join(core_path, 'core.data'), "r") as file:
            return file.read()

    def check_updated_core(self, core_data=None, force=False):
        existing_core_data = self.read_core_data(ProjectConfig.core_dir())
        if core_data and existing_core_data and core_data == existing_core_data and not force:
            print('[red]Core Already up-to-date[/]')
            return True

        FileUtils.delete(ProjectConfig.core_dir(), silence=True)
        return False

    def download_core(self, path_to_save):
        print(f"[green]Downloading core: {self.build}\nOS:{self.os}\nURL:{self.url}")
        sb.call(f"curl {self.url} --output {path_to_save}", shell=True)

    @staticmethod
    def fix_double_folder(path_to_core=ProjectConfig.core_dir()):
        path = join(path_to_core, basename(path_to_core))
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
        self.download_core(ProjectConfig.core_archive())
        FileUtils.unpacking_via_7zip(ProjectConfig.core_archive(), ProjectConfig.core_dir(), delete=True)
        self.fix_double_folder(ProjectConfig.core_dir())
        self.write_core_date_on_file(core_status.headers['Last-Modified'])
        self.change_core_access()
        self.xml.generate_doc_renderer_config(), print('[green]-' * 90)
