# -*- coding: utf-8 -*-
from os.path import basename
from rich import print
from rich.progress import track

from frameworks.StaticData import StaticData
from frameworks.decorators import highlighter
from host_tools import HostInfo, File, Shell
from telegram import Telegram

from .XmllintReport import XmllintReport


class XmlLint:
    """
    Utility class for XML linting.
    """

    def __init__(self):
        self.host = HostInfo()
        self.report = XmllintReport()
        self.xmllint_exceptions = []
        self.tmp_dir = StaticData.tmp_dir
        self.full_errors, self.parser_error = [], []
        self.ooxml_formats = (
            ".docx", ".docm", ".dotx", ".dotm",
            ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xlam",
            ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm", ".sldx", ".sldm", ".thmx"
        )

    def run_test(self, xml_path):
        """
        Run XML linting test using xmllint command.
        :param xml_path: Path to the XML file to be linted.
        :return: Output of xmllint command.
        """
        try:
            return Shell.get_output(f"xmllint --debug {xml_path} | grep 'parser error'")
        except Exception as e:
            print(f"[bold red]Exception: {e}\nwhen checking via xmllint a file_name: {xml_path}")
            self.xmllint_exceptions.append(f'Exception: {e}, file_name: {xml_path}')
            return None

    @staticmethod
    def find_parser_error(xmllint_output):
        """
        Find parser error in the xmllint output.
        :param xmllint_output: Output of xmllint command.
        :return: Parser error message if found, None otherwise.
        """
        for line in xmllint_output.split('\n'):
            if ':' in line:
                key, value = line.split(':', 1)
                return value if 'parser error' in value else key if 'parser error' in key else None

    @highlighter(color='red')
    def xml_error_handler(self, output, dir_path):
        """
        Handle XML linting errors.
        :param output: Output of xmllint command.
        :param dir_path: Directory path.
        """
        self.full_errors.append(output.replace(dir_path, ""))
        self.parser_error.append(self.find_parser_error(output))
        print(f"[red]{output}")

    def check_xml(self, dir_path, file_name, report_path, conversion_direction):
        """
        Check XML files in a directory for linting errors.
        :param dir_path: Path to the directory containing XML files.
        :param file_name: Name of the XML file.
        :param report_path: Path to the report file.
        :param conversion_direction: Conversion direction.
        """
        self.full_errors, self.parser_error = [], []
        for xml in File.get_paths(dir_path, '.xml'):
            output = self.run_test(xml)
            self.xml_error_handler(output, dir_path) if output else ...
        if self.full_errors:
            self.report.write(
                report_path, "a",
                [file_name, self.parser_error, self.full_errors, conversion_direction]
            )

    def check_ooxml_files(self, dir_path, report_path, conversion_direction):
        """
        Check OOXML files for linting errors.
        :param dir_path: Path to the directory containing OOXML files.
        :param report_path: Path to the report file.
        :param conversion_direction: Conversion direction.
        """
        for file_path in track(File.get_paths(dir_path, self.ooxml_formats), description="Xmllint checking..."):
            print(f"[green]File in test:[/] {basename(file_path)}")
            test_folder = File.unique_name(self.tmp_dir)
            File.unpacking_zip(file_path, test_folder)
            self.check_xml(test_folder, basename(file_path), report_path, conversion_direction)
            File.delete(test_folder, stdout=False)

    def run_tests(self, dir_path, conversion_direction=None):
        """
        Run XML linting tests.
        :param dir_path: Path to the directory containing XML files.
        :param conversion_direction: Conversion direction.
        :return: Path to the report file.
        """
        if self.host.os != 'windows':
            report_path = self.report.report_path()
            self.report.write(report_path, "w", ["File", "parser_error", "Error_full", "direction"])
            self.check_ooxml_files(dir_path, report_path, conversion_direction)
            print(f"[bold red]{self.report.read(report_path)}")
            Telegram.send_message(f"Exceptions: {self.xmllint_exceptions}") if self.xmllint_exceptions else ...
            print(f"[bold green]|INFO| Report created at:[/] {report_path}")
            return report_path
