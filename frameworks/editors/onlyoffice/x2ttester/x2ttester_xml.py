# -*- coding: utf-8 -*-
from os import scandir
from os.path import join, isfile, isdir
import xml.etree.cElementTree as ET

from rich import print

from .X2tTesterConfig import X2tTesterConfig
from .host_config import HostConfig

class X2tTesterXml:
    def __init__(self, config: X2tTesterConfig):
        self.config = config
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
        ET.SubElement(root, "reportPath").text = self.config.report_path
        ET.SubElement(root, "inputDirectory").text = self._delete_last_slash(self.config.input_dir)
        ET.SubElement(root, "outputDirectory").text = self._delete_last_slash(self.config.output_dir)
        ET.SubElement(root, "x2tPath").text = self._generate_x2t_path()
        ET.SubElement(root, "cores").text = str(self.config.cores)
        ET.SubElement(root, "timeout").text = str(self._generate_timeout(self.config.timeout))
        ET.SubElement(root, "troughConversion").text = self._get_parameter(self.config.trough_conversion)
        ET.SubElement(root, "errorsOnly").text = self._get_parameter(self.config.errors_only)
        ET.SubElement(root, "deleteOk").text = self._get_parameter(self.config.delete)
        ET.SubElement(root, "timestamp").text = self._get_parameter(self.config.timestamp)
        ET.SubElement(root, "saveEnvironment").text = self._get_save_environment()
        if input_format:
            ET.SubElement(root, "input").text = input_format
        if output_format:
            ET.SubElement(root, "output").text = output_format
        if files:
            ET.SubElement(root, "inputFilesList").text = files
        if self.config.fonts_dir:
            if isdir(self.config.fonts_dir) and any(scandir(self.config.fonts_dir)):
                print(f"[bold green]|INFO| Custom fonts will be generated.")
                ET.SubElement(ET.SubElement(root, "fonts", system="0"), "directory").text = self.config.fonts_dir
            else:
                print(f"[bold bright_cyan]|INFO| System fonts will be generated.")
        return root

    @staticmethod
    def _get_parameter(value: bool) -> str:
        return '1' if value else '0'

    def _get_save_environment(self) -> str:
        return '0' if self.config.environment_off else '1'

    @staticmethod
    def _delete_last_slash(path):
        return path.rstrip(path[-1]) if path[-1] in ['/', '\\'] else path

    def _generate_x2t_path(self):
        x2t_path = join(self.config.x2ttester_dir, self.host.x2t)
        if not isfile(x2t_path):
            raise Exception(f'[bold red]|ERROR| Check the existence of x2t, path: {x2t_path}')
        return x2t_path

    @staticmethod
    def _generate_timeout(value: int | None) -> int:
        if isinstance(value, int) and value > 0:
            return value
        raise Exception(
            f"[bold red]|WARNING| The `timeout` parameter to end conversion when time has elapsed will be disabled."
        )
