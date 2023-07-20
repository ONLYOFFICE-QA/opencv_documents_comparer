# -*- coding: utf-8 -*-
import subprocess as sb
from os import chdir
from os.path import join, exists

from rich import print
from frameworks.host_control.FileUtils import FileUtils
from .host_config import HostConfig
from .x2ttester_xml import X2tTesterXml


class X2tTester:
    def __init__(self, input_dir: str, output_dir: str, x2ttester_dir: str, fonts_dir: str = None):
        self.input_dir = input_dir
        self.results = output_dir
        self.x2ttester_dir = x2ttester_dir
        self.xml = X2tTesterXml(x2ttester_dir=self.x2ttester_dir, fonts_dir=fonts_dir)
        self.x2ttester_path = join(self.x2ttester_dir, HostConfig().x2ttester)

    def conversion(self, report_path: str, input_format: str, output_format: str, listxml_path: str = None) -> None:
        if not exists(self.x2ttester_path):
            raise print(f"[bold red]|ERROR| X2tTester does not exist on the path: {self.x2ttester_path}")
        chdir(self.x2ttester_dir)
        tmp_parameters_xml = self.xml.create(
            self.xml.parameters(self.input_dir, self.results, report_path, input_format, output_format, listxml_path),
            FileUtils.random_name(self.x2ttester_dir, '.xml')
        )
        sb.call(f"{self.x2ttester_path} {tmp_parameters_xml}", shell=True)
        FileUtils.delete(tmp_parameters_xml, silence=True)
