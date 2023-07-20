# -*- coding: utf-8 -*-
from os import scandir
from os.path import join, isfile, exists
import xml.etree.cElementTree as ET

from rich import print

import config
from frameworks.StaticData import StaticData
from frameworks.editors.onlyoffice.x2ttester.host_config import HostConfig

from frameworks.host_control import FileUtils


class X2tTesterXml:

    def __init__(self, x2ttester_dir: str, fonts_dir: str = None):
        self._cores = config.cores
        self.timeout = self._generate_timeout(config.timeout)
        self.errors_only = self._parameter_validator(config.errors_only)
        self.delete = self._parameter_validator(config.delete)
        self.timestamp = self._parameter_validator(config.timestamp)
        self.host = HostConfig()
        self.x2ttester_dir = x2ttester_dir
        self.fonts_dir = fonts_dir

    @staticmethod
    def files_list(files: list) -> ET.Element:
        root = ET.Element("files")
        for file in files:
            ET.SubElement(root, "file_name").text = file
        return root

    def parameters(self,
                   input_dir: str,
                   result_dir: str,
                   report_path: str,
                   input_format: str = None,
                   output_format: str = None,
                   files: list = None
                   ) -> ET.Element:

        root = ET.Element("Settings")
        ET.SubElement(root, "reportPath").text = report_path
        ET.SubElement(root, "inputDirectory").text = FileUtils.delete_last_slash(input_dir)
        ET.SubElement(root, "outputDirectory").text = FileUtils.delete_last_slash(result_dir)
        ET.SubElement(root, "x2tPath").text = self._generate_x2t_path()
        ET.SubElement(root, "cores").text = self._cores
        ET.SubElement(root, "timeout").text = self.timeout
        if input_format:
            ET.SubElement(root, "input").text = input_format
        if output_format:
            ET.SubElement(root, "output").text = output_format
        if self.errors_only:
            ET.SubElement(root, "errorsOnly").text = self.errors_only
        if self.delete:
            ET.SubElement(root, "deleteOk").text = self.delete
        if self.timestamp:
            ET.SubElement(root, "timestamp").text = self.timestamp
        if files:
            ET.SubElement(root, "inputFilesList").text = files
        if self.fonts_dir and exists(self.fonts_dir) and any(scandir(self.fonts_dir)):
            print(f"[bold green]|INFO| Custom fonts will be generated.")
            ET.SubElement(ET.SubElement(root, "fonts", system="0"), "directory").text = self.fonts_dir
        else:
            print(f"[bold bright_cyan]|INFO| System fonts will be generated.")
        return root

    @staticmethod
    def _parameter_validator(value: bool) -> str:
        if isinstance(value, bool):
            return '1' if value else '0'
        raise print("[bold red]|ERROR| Parameters: 'errors_only', 'delete', 'timestamp' must be True or False")

    def _generate_x2t_path(self):
        x2t_path = join(self.x2ttester_dir, self.host.x2t)
        if not isfile(x2t_path):
            raise print(f'[bold red]|ERROR| Check the existence of x2t, path: {x2t_path}')
        return x2t_path

    @property
    def _cores(self):
        return self.__cores

    @_cores.setter
    def _cores(self, cores_num: int):
        if not isinstance(cores_num, int) and cores_num <= 0:
            raise print('[bold red]|ERROR| The number of cores must be integer and greater than 0')
        self.__cores = str(cores_num)

    @staticmethod
    def _generate_timeout(value: int | None) -> str:
        if value:
            if isinstance(value, int):
                return str(value)
            raise print('[bold red]|ERROR| Parameter `timeout` in config mast be integer.')
        input(
            f"[bold red]|WARNING| The `timeout` parameter to end conversion when time has elapsed will be disabled. "
            f"Press `Enter` to continue."
        )
        return '0'

    @staticmethod
    def create(xml: ET.Element, xml_path: str) -> str:
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        tree.write(xml_path, encoding="UTF-8", xml_declaration=True)
        return xml_path
