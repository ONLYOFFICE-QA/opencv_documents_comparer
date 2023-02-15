# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import join, basename

from rich import print
from rich.progress import track

from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions
from framework.actions.report_actions import ReportActions
from framework.actions.x2t_actions import X2t
from framework.actions.xml_actions import XmlActions
from framework.singleton import singleton
from framework.telegram import Telegram


@singleton
class XmlLint:
    def __init__(self):
        self.host = HostActions()
        self.csv = ReportActions()
        self.time_pattern = f"{datetime.now().strftime('%H_%M_%S')}"
        self.xmllint_exceptions = []
        self.ooxml_formats = (".docx", ".docm", ".dotx", ".dotm",
                              ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam",
                              ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx")

    def generate_xmllint_report_path(self):
        x2t_version = X2t.x2t_version()
        major_version = XmlActions.generate_major_version(x2t_version)
        report_dir = join(ProjectConfig.reports_dir(), major_version, self.host.os, f"xmllint")
        FileUtils.create_dir(report_dir)
        return join(report_dir, f"{x2t_version}_{self.time_pattern}.csv")

    def run_xmllint_test(self, path_to_xml_file):
        try:
            return FileUtils.output_cmd(f"xmllint --debug {path_to_xml_file} | grep 'parser error'")
        except Exception as e:
            print(f"[bold red]Exception: {e}\nwhen checking via xmllint a file: {path_to_xml_file}")
            self.xmllint_exceptions.append(f'Exception: {e}, file: {path_to_xml_file}')
            return None

    @staticmethod
    def find_error_in_output(xmllint_output):
        for line in xmllint_output.split('\n'):
            if ':' in line:
                key, value = line.split(':', 1)
                return value if 'parser error' in value else key if 'parser error' in key else None

    def check_xml(self, dir_path, file, xmllint_report, direction=None):
        xml_errors, parser_error = [], []
        for xml in FileUtils.get_file_paths(dir_path, '.xml'):
            output = self.run_xmllint_test(xml)
            if output:
                xml_errors.append(output.replace(dir_path, "")), parser_error.append(self.find_error_in_output(output))
                print(f"{'[red]-' * 90}\n[red]{output}\n{'[red]-' * 90}")
        self.csv.csv_writer(xmllint_report, "a", [file, parser_error, xml_errors, direction]) if xml_errors else ...

    def check_ooxml_files(self, dir_path, xmllint_report, conversion_direction=None):
        file_paths_array = FileUtils.get_file_paths(dir_path, self.ooxml_formats)
        for file_path in track(file_paths_array, description="Check files via xmllint..."):
            print(f"[green]File in test:[/] {basename(file_path)}")
            xmllint_test_folder = FileUtils.random_name(ProjectConfig.TMP_DIR)
            FileUtils.unpacking_via_zip_file(file_path, xmllint_test_folder)
            self.check_xml(xmllint_test_folder, basename(file_path), xmllint_report, conversion_direction)
            FileUtils.delete(xmllint_test_folder, silence=True)

    def run_tests(self, dir_path, conversion_direction=None):
        if self.host.os != 'windows':
            report = self.generate_xmllint_report_path()
            self.csv.csv_writer(report, "w", ["File", "parser_error", "Error_full", "direction"])
            self.check_ooxml_files(dir_path, report, conversion_direction)
            print(f"[bold red]{self.csv.read_csv_via_pandas(report)}")
            Telegram.send_message(f"Exceptions: {self.xmllint_exceptions}") if self.xmllint_exceptions else ...
            print(f"[bold green]|INFO| Report created at:[/] {report}")
            return report
