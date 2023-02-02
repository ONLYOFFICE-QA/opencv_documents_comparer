# -*- coding: utf-8 -*-
import json
import subprocess as sb
from os import chdir
from os.path import join

from rich import print
from rich.prompt import Prompt

import settings
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.document_actions import DocActions
from framework.actions.host_actions import HostActions
from framework.actions.x2t_actions import X2t
from framework.actions.xml_actions import XmlActions
from framework.singleton import singleton
from framework.xmllint import XmlLint


@singleton
class Converter:
    def __init__(self):
        self.xml = XmlActions()
        self.host = HostActions()
        self.xmllint = XmlLint()
        self.x2t_version = None

    @staticmethod
    def getting_formats(direction=None):
        _direction = direction if direction else Prompt.ask('Input formats with -', default=None, show_default=False)
        if _direction:
            if '-' in _direction:
                return _direction.split('-')[0], _direction.split('-')[1]
            return None, _direction
        return None, None

    def conversion_via_x2ttester(self, input_format, output_format, ls=False):
        chdir(ProjectConfig.core_dir())
        self.x2t_version = X2t.x2t_version()
        print(f"[bold green]|INFO| The conversion is running on x2t version: [bold red]{self.x2t_version}")
        path_to_list = self.xml.generate_files_list() if ls else None
        tmp_xml = self.xml.generate_x2ttester_parameters(input_format, output_format, path_to_list, self.x2t_version)
        sb.call(f"{join(ProjectConfig.core_dir(), self.host.x2ttester)} {tmp_xml}", shell=True)
        print(f"[bold red]|INFO| x2t version: {self.x2t_version}")

    def convert_from_extension_array(self):
        for extensions in json.load(open(f"{ProjectConfig.PROJECT_DIR}/data/extension_array.json")).items():
            src_ext = extensions[0]
            for cnv_ext in extensions[1]:
                settings.delete = '0'
                self.conversion_via_x2ttester(src_ext, cnv_ext)
                result_dir = join(ProjectConfig.result_dir(), f"{self.X2ttester.x2t_version}_{src_ext}_{cnv_ext}")
                DocActions.copy_result_x2ttester(result_dir, cnv_ext)
                self.xmllint.run_tests(ProjectConfig.tmp_result_dir())
                FileUtils.delete(ProjectConfig.tmp_result_dir())
