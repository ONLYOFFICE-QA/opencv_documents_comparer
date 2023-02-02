# -*- coding: utf-8 -*-
import subprocess as sb
from os import chdir

from rich import print

from configurations.project_configurator import ProjectConfig
from framework.actions.host_actions import HostActions
from framework.actions.x2t_actions import X2t
from framework.actions.xml_actions import XmlActions


class X2ttester:
    def __init__(self):
        self.host_config = HostActions()
        self.xml = XmlActions()
        self.x2t_version = None

    def run_x2ttester(self, input_format, output_format, path_to_list=None):
        self.x2t_version = X2t.x2t_version()
        print(f"[bold green]|INFO|The conversion is running on x2t version: [red]{self.x2t_version}")
        tmp_xml = self.xml.generate_x2ttester_parameters(input_format, output_format, path_to_list, self.x2t_version)
        chdir(ProjectConfig.core_dir())
        sb.call(f"{ProjectConfig.core_dir()}/{self.host_config.x2ttester} {tmp_xml}", shell=True)
        print(f"[bold red]|WARNING|x2t version: {self.x2t_version}")
