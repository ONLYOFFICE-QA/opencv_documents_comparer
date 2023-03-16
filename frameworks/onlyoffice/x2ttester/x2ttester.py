# -*- coding: utf-8 -*-
import json
import subprocess as sb
from os import chdir, walk
from os.path import join, dirname, basename, realpath

from rich import print
from rich.prompt import Prompt

import settings
from frameworks.StaticData import StaticData
from frameworks.host_control.FileUtils import FileUtils
from frameworks.onlyoffice.x2t.x2t import X2t
from frameworks.onlyoffice.x2ttester.host_config import HostConfig
from frameworks.xmllint import XmlLint
from frameworks.onlyoffice.x2ttester.x2ttester_report import X2ttesterReport
from ..xml.x2ttester_xml import X2tTesterXml


class X2ttester:
    def __init__(self):
        self.tmp_dir = join(StaticData.TMP_DIR, 'cnv')
        self.x2ttester_dir = StaticData.core_dir()
        self.result_dir = StaticData.result_dir()
        self.x2ttester_path = join(self.x2ttester_dir, HostConfig().x2ttester)
        self.x2t_version = X2t.version(self.x2ttester_dir)
        self.xmllint = XmlLint()
        self.xml = X2tTesterXml(self.tmp_dir)
        self.report = X2ttesterReport()

    def conversion_from_files_list(self, input_format, output_format):
        list_xml_path = self.xml.create_tmp_files_list(settings.files_array)
        report_path = self.conversion(input_format, output_format, list_xml_path)
        FileUtils.delete(list_xml_path)
        return report_path

    def copy_result(self, path_to, output_format, delete=False):
        FileUtils.create_dir(path_to)
        if output_format in ["png", "jpg"]:
            for dir_path in FileUtils.get_dir_paths(self.tmp_dir, f".{output_format}"):
                FileUtils.copy(dir_path, join(path_to, basename(dir_path)), silence=True)
        else:
            for file_path in FileUtils.get_paths(self.tmp_dir, f".{output_format}"):
                FileUtils.copy(file_path, join(path_to, basename(file_path)), silence=True)
        FileUtils.delete(self.tmp_dir) if delete else ...

    # @param [Boolean] ls enables/disables XML generation with the file_name names to be converted
    # @param [String] output_format
    # @param [String] input_format
    # @return Path to a x2ttester report.
    def conversion(self, input_format: str, output_format: str, files_list_path: str = None) -> str:
        print(f"[bold green]|INFO| The conversion is running on x2t version: [bold red]{self.x2t_version}")
        chdir(self.x2ttester_dir)
        report_path = X2ttesterReport().random_report_path(self.x2t_version)
        tmp_xml = self.xml.create_tmp_parameters(input_format, output_format, files_list_path, report_path)
        sb.call(f"{self.x2ttester_path} {tmp_xml}", shell=True)
        print(f"[bold red]\n|INFO|{'-' * 90}\nx2t version: {self.x2t_version}\n{'-' * 90}")
        return FileUtils.last_modified_file(dirname(report_path))

    # @param [String] direction X2ttester direction
    @staticmethod
    def getting_formats(direction: str = None):
        _direction = direction if direction else Prompt.ask('Input formats with -', default=None, show_default=False)
        if _direction:
            if '-' in _direction:
                return _direction.split('-')[0], _direction.split('-')[1]
            return None, _direction
        return None, None

    def convert_from_extension_array(self):
        xmllint_report, x2ttester_reports, = None, []
        for extensions in json.load(open(f"{dirname(realpath(__file__))}/assets/extension_array.json")).items():
            src_ext = extensions[0]
            for cnv_ext in extensions[1]:
                x2ttester_reports.append(self.conversion(src_ext, cnv_ext))
                result_dir = join(self.result_dir, f"{self.x2t_version}_{src_ext}_{cnv_ext}")
                self.copy_result(result_dir, cnv_ext)
                xmllint_report = self.xmllint.run_tests(self.tmp_dir, f"{src_ext}-{cnv_ext}")
                FileUtils.delete(self.tmp_dir)
        return xmllint_report, self.report.merge_x2ttester_reports(x2ttester_reports, self.x2t_version)
