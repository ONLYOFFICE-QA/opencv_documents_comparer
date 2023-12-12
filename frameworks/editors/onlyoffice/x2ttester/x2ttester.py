# -*- coding: utf-8 -*-
import subprocess as sb
from os import chdir
from os.path import join, exists

from rich import print
from host_tools import File

from .Data import Data
from .host_config import HostConfig
from .x2ttester_xml import X2tTesterXml


class X2tTester:
    def __init__(self, data: Data):
        self.data = data
        self.x2ttester_dir = data.x2ttester_dir
        self.x2ttester_path = join(self.x2ttester_dir, HostConfig().x2ttester)
        self.xml = X2tTesterXml(data)

    def conversion(self, input_format: str, output_format: str, listxml_path: str = None) -> None:
        if not exists(self.x2ttester_path):
            raise print(f"[bold red]|ERROR| X2tTester does not exist on the path: {self.x2ttester_path}")
        chdir(self.x2ttester_dir)
        tmp_parameters_xml = self.xml.create(
            self.xml.parameters(input_format, output_format, listxml_path),
            File.unique_name(self.x2ttester_dir, '.xml')
        )
        sb.call(f"{self.x2ttester_path} {tmp_parameters_xml}", shell=True)
        File.delete(tmp_parameters_xml, stdout=False)
