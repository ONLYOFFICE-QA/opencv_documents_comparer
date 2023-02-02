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
from framework.singleton import singleton
from framework.telegram import Telegram


@singleton
class XmlLint:
    def __init__(self):
        self.host = HostActions()
        self.csv = ReportActions()
        self.xmllint_error_list = []
        self.report_dir = ProjectConfig.xmllint_report_dir()
        self.report_name = f"{datetime.now().strftime('%H_%M_%S')}.csv"

    def create_xmllint_report(self):
        FileUtils.create_dir(self.report_dir)
        self.report_name = f"{X2t.x2t_version()}_{self.report_name}"
        self.csv.csv_writer(join(self.report_dir, self.report_name), "w", ["File", "Path", "Error", "Exit_Code"])

    @staticmethod
    def run_xmllint_test(path_to_xml_file):
        try:
            return FileUtils.output_cmd(f"xmllint --debug {path_to_xml_file} | grep 'parser error'")
        except Exception as e:
            print(f"[bold red]Exception: {e}\nwhen checking via xmllint a file: {path_to_xml_file}")
            Telegram.send_message(f'Exception: {e}\nwhen checking via xmllint a file: {path_to_xml_file}')
            return None

    def check_via_xmllint(self, path_to_xml_folder, file_name):
        for xml in FileUtils.get_files_by_extensions(path_to_xml_folder, '.xml'):
            output = self.run_xmllint_test(xml)
            if output:
                self.xmllint_error_list.append(file_name)
                self.csv.csv_writer(join(self.report_dir, self.report_name), "a", [file_name, xml, output, '1'])
                print(f"{'[red]-' * 90}\n[red]{output}\n{'[red]-' * 90}")

    def run_tests(self, path_to_files):
        if self.host.os != 'windows':
            self.create_xmllint_report()
            files_array = FileUtils.get_files_by_extensions(path_to_files, (".xlsx", ".pptx", ".docx"))
            for file_path in track(files_array, description="Check files via xmllint..."):
                print(f"[green]File in test:[/] {basename(file_path)}")
                xmllint_test_folder = FileUtils.random_name(ProjectConfig.TMP_DIR)
                FileUtils.unpacking_via_zip_file(file_path, xmllint_test_folder)
                self.check_via_xmllint(xmllint_test_folder, basename(file_path))
                FileUtils.delete(xmllint_test_folder, silence=True)
            print(f"[bold red]{set(self.xmllint_error_list)}") if self.xmllint_error_list else None
