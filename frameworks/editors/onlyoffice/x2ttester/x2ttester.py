# -*- coding: utf-8 -*-
import subprocess as sb
from os import chdir
from os.path import join, exists

from host_tools import File
from .X2tTesterConfig import X2tTesterConfig
from .x2ttester_xml import X2tTesterXml
from .host_config import HostConfig


class X2tTester:
    def __init__(self, config: X2tTesterConfig):
        self.x2ttester_dir = config.x2ttester_dir
        self.x2ttester_path = join(self.x2ttester_dir, HostConfig().x2ttester)
        self.xml = X2tTesterXml(config)

    def conversion(self, input_format: str, output_format: str, listxml_path: str = None) -> None:
        """
        Method for performing file conversion.
        :param input_format: Input file format.
        :param output_format: Output file format.
        :param listxml_path: Path to the list.xml file, defaults to None.
        """
        self.check_x2ttester_exists()
        chdir(self.x2ttester_dir)
        param_xml = self.create_param_xml(input_format, output_format, listxml_path)
        sb.call(f"{self.x2ttester_path} {param_xml}", shell=True)
        File.delete(param_xml, stdout=False)

    def check_x2ttester_exists(self) -> None:
        if not exists(self.x2ttester_path):
            raise FileNotFoundError(f"[bold red]|ERROR| X2tTester does not exist on the path: {self.x2ttester_path}")

    def create_param_xml(self, input_format: str, output_format: str, listxml_path: str = None) -> str:
        return self.xml.create(
            self.xml.parameters(input_format, output_format, listxml_path),
            File.unique_name(self.x2ttester_dir, '.xml')
        )
