# -*- coding: utf-8 -*-
from os import chdir, sep
from os.path import join, isfile, exists, abspath, normcase
from time import sleep

import psutil
from rich import print

from frameworks.decorators.decorators import highlighter
from frameworks.StaticData import StaticData
from host_tools import File, HostInfo

from .x2t_libs_xml import X2tLibsXML
from .UrlGenerator import UrlGenerator


class Core:

    def __init__(self, version: str = None):
        """
        Initializes the Core object.
        :param version: Version of the core.
        """
        self.os = HostInfo().os
        self.url = UrlGenerator(version).url
        self.version = version
        self.core_dir = StaticData.core_dir()
        self.tmp_dir = StaticData.tmp_dir
        self.project_dir = StaticData.project_dir
        self.data_file = join(self.core_dir, 'core.data')
        self.xml = X2tLibsXML()

    @highlighter(color='green')
    def getting(self, force: bool = False) -> None:
        """
        :param force: Whether to force update the core files. Defaults to False.
        """
        self._delete_core_dir() if force else ...
        headers = File.get_headers(self.url)
        if not headers or self._check_updated_core(core_data=headers['Last-Modified']):
            return
        self._delete_core_dir()
        self._download()
        File.unpacking_7z(join(self.tmp_dir, "core.7z"), self.core_dir, delete_archive=True)
        File.fix_double_dir(self.core_dir)
        File.change_access(self.core_dir)
        File.write(self.data_file, headers['Last-Modified'], mode='w')
        self.xml.create_doc_renderer_config()

    def _read_core_data(self) -> str | None:
        """
        Reads core data from the data file.
        :return: Core data string if file exists, None otherwise.
        """
        if not isfile(self.data_file):
            return None
        return File.read(self.data_file, mode='r')

    def _check_updated_core(self, core_data: str = None) -> bool:
        """
        Checks if the core is already updated.
        :param core_data: Last modified date of the core.
        :return: True if core is up-to-date, False otherwise.
        """
        existing_core_data = self._read_core_data()
        if core_data and existing_core_data and core_data == existing_core_data:
            print('[red]|INFO| Core Already up-to-date[/]')
            return True
        return False

    def _download(self) -> None:
        """
        Downloads the core files.
        """
        print(f"[green]|INFO| Downloading core\nVersion: {self.version}\nOS: {self.os}\nURL: {self.url}")
        File.download(self.url, self.tmp_dir, "core.7z")

    def _delete_core_dir(self) -> None:
        """
        Deletes the core directory.
        """
        chdir(self.project_dir)
        self._kill_core_processes()

        for _ in range(5):
            File.delete(self.core_dir, stdout=False, stderr=True)
            if not exists(self.core_dir):
                return
            sleep(1)

        raise PermissionError(
            f"|ERROR| Core directory is locked and could not be deleted: {self.core_dir}. "
            "Close the processes using files in it (x2t, x2ttester, antivirus) and try again."
        )

    def _kill_core_processes(self) -> None:
        """
        Kills the processes that lock the core directory (their executable or working directory
        is inside it) together with their children processes.
        """
        killed = []

        for process in psutil.process_iter():
            if not self._is_core_process(process):
                continue
            try:
                children_processes = process.children(recursive=True)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                children_processes = []

            for _process in [process, *children_processes]:
                try:
                    name = _process.name()
                    _process.kill()
                    killed.append(_process)
                    print(f"[red]|INFO| Killed process: {name}, pid: {_process.pid}")
                except psutil.NoSuchProcess:
                    pass
                except psutil.AccessDenied:
                    print(f"[bold red]|ERROR| AccessDenied when killing process pid: {_process.pid}")

        psutil.wait_procs(killed, timeout=5)

    def _is_core_process(self, process: psutil.Process) -> bool:
        """
        Checks whether the process executable or working directory is inside the core directory.
        :param process: Process to check.
        :return: True if the process locks the core directory, False otherwise.
        """
        core_dir = normcase(abspath(self.core_dir))
        try:
            paths = [process.exe(), process.cwd()]
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            return False

        for path in paths:
            if not path:
                continue
            path = normcase(abspath(path))
            if path == core_dir or path.startswith(core_dir + sep):
                return True
        return False
