# -*- coding: utf-8 -*-
import json
import time

from rich import print

import settings
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.document_actions import DocActions
from framework.actions.x2t_actions import X2t
from framework.actions.x2ttester_actions import X2ttester
from framework.actions.xml_actions import XmlActions
from framework.singleton import singleton
from framework.xmllint import XmlLint


@singleton
class Converter:
    def __init__(self):
        self.xml = XmlActions()
        self.x2ttester = X2ttester()
        self.xmllint = XmlLint()

    @staticmethod
    def getting_formats(direction):
        if direction:
            return direction.split('-')[0], direction.split('-')[1]
        else:
            return None, None

    def conversion_via_x2ttester(self, input_format, output_format, ls=False):
        start_time = time.time()
        if ls:
            list_xml = self.xml.generate_files_list()
            self.x2ttester.run_x2ttester(input_format, output_format, list_xml)
        else:
            self.x2ttester.run_x2ttester(input_format, output_format)
        print(f"[bold red]{X2t.x2t_version()}")
        print(time.time() - start_time)

    def convert_from_extension_array(self):
        for extensions in json.load(open(f"{ProjectConfig.PROJECT_DIR}/data/extension_array.json")).items():
            src_extension = extensions[0]
            for cnv_extension in extensions[1]:
                settings.delete = '0'
                self.x2ttester.run_x2ttester(src_extension, cnv_extension)
                result_folder = f"{ProjectConfig.result_dir()}/{X2t.x2t_version()}_{src_extension}_{cnv_extension}"
                DocActions.copy_result_x2ttester(result_folder, cnv_extension)
                self.xmllint.run_tests(ProjectConfig.tmp_result_dir())
                FileUtils.delete(ProjectConfig.tmp_result_dir())
