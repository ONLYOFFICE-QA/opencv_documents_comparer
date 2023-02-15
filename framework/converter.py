# -*- coding: utf-8 -*-
import json
import subprocess as sb
from datetime import datetime
from os import chdir
from os.path import join

from rich import print
from rich.prompt import Prompt

from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.document_actions import DocActions
from framework.actions.host_actions import HostActions
from framework.actions.report_actions import ReportActions
from framework.actions.x2t_actions import X2t
from framework.actions.xml_actions import XmlActions
from framework.singleton import singleton
from framework.xmllint import XmlLint


@singleton
class Converter:
    def __init__(self):
        self.report = ReportActions()
        self.xml = XmlActions()
        self.host = HostActions()
        self.xmllint = XmlLint()
        self.x2t_version = None

    # @param [String] direction Conversion direction
    @staticmethod
    def getting_formats(direction=None):
        _direction = direction if direction else Prompt.ask('Input formats with -', default=None, show_default=False)
        if _direction:
            if '-' in _direction:
                return _direction.split('-')[0], _direction.split('-')[1]
            return None, _direction
        return None, None

    # @param [Boolean] ls enables/disables XML generation with the file names to be converted
    # @param [String] output_format
    # @param [String] input_format
    # @return Path to a x2ttester report.
    def conversion_via_x2ttester(self, input_format, output_format, ls=False):
        chdir(ProjectConfig.core_dir())
        self.x2t_version = X2t.x2t_version()
        print(f"[bold green]|INFO| The conversion is running on x2t version: [bold red]{self.x2t_version}")
        path_to_list = self.xml.generate_files_list() if ls else None
        report_path, tmp_report_dir = self.xml.generate_report_paths(input_format, output_format, self.x2t_version)
        tmp_xml = self.xml.generate_x2ttester_parameters(input_format, output_format, path_to_list, report_path)
        sb.call(f"{join(ProjectConfig.core_dir(), self.host.x2ttester)} {tmp_xml}", shell=True)
        print(f"[bold red]|INFO|{'-' * 90}\nx2t version: {self.x2t_version}\n{'-' * 90}")
        return FileUtils.last_modified_file(tmp_report_dir)

    # Merges several x2ttester reports into one.
    # @param [Array] x2ttester_reports Array with paths to x2ttester reports
    # @return Path to a x2ttester merged report.
    def merge_x2ttester_reports(self, x2ttester_reports):
        x2ttester_merged_report_name = f"{self.x2t_version}_{datetime.now().strftime('%H_%M_%S')}_array.csv"
        x2ttester_merged_report = join(self.xml.generate_report_dir(self.x2t_version), x2ttester_merged_report_name)
        self.report.merge_reports(x2ttester_reports, x2ttester_merged_report)
        return x2ttester_merged_report

    def convert_from_extension_array(self):
        xmllint_report, x2ttester_reports, = None, []
        for extensions in json.load(open(f"{ProjectConfig.PROJECT_DIR}/data/extension_array.json")).items():
            src_ext = extensions[0]
            for cnv_ext in extensions[1]:
                x2ttester_reports.append(self.conversion_via_x2ttester(src_ext, cnv_ext))
                self.x2t_version = X2t.x2t_version()
                result_dir = join(ProjectConfig.result_dir(), f"{self.x2t_version}_{src_ext}_{cnv_ext}")
                DocActions.copy_result_x2ttester(result_dir, cnv_ext)
                xmllint_report = self.xmllint.run_tests(ProjectConfig.tmp_result_dir(), f"{src_ext}-{cnv_ext}")
                FileUtils.delete(ProjectConfig.tmp_result_dir())
        return xmllint_report, self.merge_x2ttester_reports(x2ttester_reports)
