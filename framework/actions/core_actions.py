# -*- coding: utf-8 -*-
import re
import subprocess as sb
from os import chdir, listdir, scandir
from os.path import join, isdir, isfile

import requests
from rich import print
from rich.progress import track

import settings
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions
from framework.actions.xml_actions import XmlActions


class CoreActions:
    def __init__(self, version=settings.version):
        self.host_config = HostActions()
        self.xml = XmlActions()
        self.config_version = version if version else input("Please enter version core: ")
        self.branch = self.generate_branch()
        self.os = self.host_config.os
        self.arch = self.host_config.arch
        self.build = self.generate_build()
        self.version = self.generate_version()
        self.url = self.generate_url()

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
        return f"{host}/{self.os}/core/{self.branch}/{self.version}/{self.build}/{self.arch}/core.7z"

    def check_core_on_server(self):
        branch_array = ["release", "hotfix"]
        branch_array.remove(self.branch) if self.branch in branch_array else None
        branch_array.insert(0, self.branch)
        for branch in branch_array:
            if self.branch != branch:
                self.branch = branch
                self.url = self.generate_url()
            core_status = requests.head(self.url)
            if core_status.status_code == 200:
                return core_status

        raise print(f"[bold red]Core not found, check version and branch settings\n"
                    f"URL: {self.url}\nCurl response:\n{requests.head(self.url).status_code}[/]")

    def change_core_access(self):
        if self.os == 'windows':
            return

        sb.call(f'chmod +x {ProjectConfig.core_dir()}/*', shell=True)

    def generate_allfonts(self):
        chdir(ProjectConfig.core_dir())
        sb.call(join(ProjectConfig.core_dir(), self.host_config.standardtester))
        chdir(ProjectConfig.PROJECT_DIR)

    @staticmethod
    def getting_core_date(core_status):
        for line in core_status.split("\n"):
            if ':' not in line:
                continue
            key, value = line.strip().split(':', 1)
            if key.upper() == "LAST-MODIFIED":
                return value
        return ''

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
            raise print('[red]Core Already up-to-date[/]')

        FileUtils.delete(ProjectConfig.core_dir(), silence=True)

    def download_core(self):
        print(f"[green]Downloading core {self.branch}/{self.version}:{self.build}\nOS: {self.os}\nURL: {self.url}")
        sb.call(f"curl {self.url} --output {ProjectConfig.core_archive()}", shell=True)

    @staticmethod
    def fix_double_core_folder(path_to_core=ProjectConfig.core_dir()):
        if isdir(join(path_to_core, 'core')):
            print('[red]Double core folder')
            for file in track(listdir(join(ProjectConfig.core_dir(), 'core')), description='[green]Fixing...'):
                FileUtils.move(join(ProjectConfig.core_dir(), 'core', file), join(ProjectConfig.core_dir(), file))
            if not any(scandir(join(ProjectConfig.core_dir(), 'core'))):
                return FileUtils.delete(join(ProjectConfig.core_dir(), 'core'))
            print("[red]Not all objects are moved")

    def copy_x2ttester_tools(self, path_from, path_to):
        FileUtils.create_dir(path_to)
        FileUtils.delete(join(path_to, self.host_config.x2ttester))
        FileUtils.copy(join(path_from, self.host_config.x2ttester), join(path_to, self.host_config.x2ttester))

    def update_x2ttester_tools(self):
        print('[green]-' * 90)
        core_dir, tools = join(ProjectConfig.TMP_DIR, 'core'), join(ProjectConfig.tools_dir(), self.host_config.os)
        self.check_core_on_server()
        FileUtils.delete(core_dir, silence=True)
        self.download_core()
        FileUtils.unpacking_via_7zip(ProjectConfig.core_archive(), core_dir, delete=True)
        self.fix_double_core_folder(core_dir)
        self.copy_x2ttester_tools(core_dir, tools)
        FileUtils.delete(core_dir)
        print('[green]-' * 90)

    def getting_core(self, force=False, copy_tools=False):
        print('[green]-' * 90)
        core_status = self.check_core_on_server()
        self.check_updated_core(core_data=core_status.headers['Last-Modified'], force=force)
        self.download_core()
        FileUtils.unpacking_via_7zip(ProjectConfig.core_archive(), ProjectConfig.core_dir(), delete=True)
        self.fix_double_core_folder(ProjectConfig.core_dir())
        if copy_tools:
            self.copy_x2ttester_tools(join(ProjectConfig.tools_dir(), self.os), ProjectConfig.core_dir())
            print(f"[blue]x2t tools copied from: {ProjectConfig.tools_dir()}/{self.os}")
        self.write_core_date_on_file(core_status.headers['Last-Modified'])
        self.change_core_access()
        self.xml.generate_doc_renderer_config()
        print('[green]-' * 90)
