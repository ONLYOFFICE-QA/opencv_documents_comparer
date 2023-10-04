# -*- coding: utf-8 -*-
from os import scandir
from os.path import join, isfile, exists
import xml.etree.cElementTree as ET

from rich import print

from .Data import Data
from .host_config import HostConfig


class X2tTesterXml:
    def __init__(self, data: Data):
        self.data = data
        self.host = HostConfig()

    @staticmethod
    def create(xml: ET.Element, xml_path: str) -> str:
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        tree.write(xml_path, encoding="UTF-8", xml_declaration=True)
        return xml_path

    @staticmethod
    def files_list(files: list) -> ET.Element:
        root = ET.Element("files")
        for file in files:
            ET.SubElement(root, "file_name").text = file
        return root

    def parameters(self, input_format: str = None, output_format: str = None, files: list = None) -> ET.Element:
        root = ET.Element("Settings")
        ET.SubElement(root, "reportPath").text = self.data.report_path
        ET.SubElement(root, "inputDirectory").text = self._delete_last_slash(self.data.input_dir)
        ET.SubElement(root, "outputDirectory").text = self._delete_last_slash(self.data.output_dir)
        ET.SubElement(root, "x2tPath").text = self._generate_x2t_path()
        ET.SubElement(root, "cores").text = str(self._get_num_cores(self.data.cores))
        ET.SubElement(root, "timeout").text = str(self._generate_timeout(self.data.timeout))
        ET.SubElement(root, "troughConversion").text = self._get_parameter(self.data.trough_conversion)
        ET.SubElement(root, "errorsOnly").text = self._get_parameter(self.data.errors_only)
        ET.SubElement(root, "deleteOk").text = self._get_parameter(self.data.delete)
        ET.SubElement(root, "timestamp").text = self._get_parameter(self.data.timestamp)
        ET.SubElement(root, "saveEnvironment").text = self._get_save_environment()
        if input_format:
            ET.SubElement(root, "input").text = input_format
        if output_format:
            ET.SubElement(root, "output").text = output_format
        if files:
            ET.SubElement(root, "inputFilesList").text = files
        if self.data.fonts_dir:
            if exists(self.data.fonts_dir) and any(scandir(self.data.fonts_dir)):
                print(f"[bold green]|INFO| Custom fonts will be generated.")
                ET.SubElement(ET.SubElement(root, "fonts", system="0"), "directory").text = self.data.fonts_dir
            else:
                print(f"[bold bright_cyan]|INFO| System fonts will be generated.")
        return root

    @staticmethod
    def _get_parameter(value: bool) -> str:
        if isinstance(value, bool):
            return '1' if value else '0'

    def _get_save_environment(self) -> str:
        return '0' if self.data.environment_off else '1'

    @staticmethod
    def _delete_last_slash(path):
        return path.rstrip(path[-1]) if path[-1] in ['/', '\\'] else path

    def _generate_x2t_path(self):
        x2t_path = join(self.data.x2ttester_dir, self.host.x2t)
        if not isfile(x2t_path):
            raise print(f'[bold red]|ERROR| Check the existence of x2t, path: {x2t_path}')
        return x2t_path

    @staticmethod
    def _get_num_cores(cores_num: int) -> int:
        if not isinstance(cores_num, int) and cores_num <= 0:
            raise print('[bold red]|ERROR| The number of cores must be integer and greater than 0')
        return cores_num

    @staticmethod
    def _generate_timeout(value: int | None) -> int:
        if isinstance(value, int) and value > 0:
            return value
        input(
            f"[bold red]|WARNING| The `timeout` parameter to end conversion when time has elapsed will be disabled. "
            f"Press `Enter` to continue."
        )
        return 0
