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
        """
        Initializes the X2tTester object.
        :param config: Configuration object for X2tTester.
        """
        self.x2ttester_dir = config.core_dir
        self.x2ttester_path = join(self.x2ttester_dir, HostConfig().x2ttester)
        self.xml = X2tTesterXml(config)

    def conversion(self, input_format: str, output_format: str, listxml_path: str = None) -> None:
        """
        Performs file conversion using X2tTester.
        :param input_format: Input file format.
        :param output_format: Output file format.
        :param listxml_path: Path to the list.xml file. Defaults to None.
        """
        self.check_x2ttester_exists()
        chdir(self.x2ttester_dir)
        param_xml = self.create_param_xml(input_format, output_format, listxml_path)
        sb.call(f"{self.x2ttester_path} {param_xml}", shell=True)
        File.delete(param_xml, stdout=False)

    def check_x2ttester_exists(self) -> None:
        """
        Checks if X2tTester exists on the specified path.
        :raises FileNotFoundError: If X2tTester is not found.
        """
        if not exists(self.x2ttester_path):
            raise FileNotFoundError(f"[bold red]|ERROR| X2tTester does not exist on the path: {self.x2ttester_path}")

    def create_param_xml(self, input_format: str, output_format: str, listxml_path: str = None) -> str:
        """
        Creates parameter XML file for conversion.
        :param input_format: Input file format.
        :param output_format: Output file format.
        :param listxml_path: Path to the list.xml file. Defaults to None.
        :returns: Path to the created XML file.
        """
        return self.xml.create(
            self.xml.parameters(input_format, output_format, listxml_path),
            File.unique_name(self.x2ttester_dir, '.xml')
        )
