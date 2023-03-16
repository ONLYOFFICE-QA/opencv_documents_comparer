# -*- coding: utf-8 -*-
from os import scandir
from os.path import join, isfile, exists

from rich import print

import settings
from frameworks.StaticData import StaticData
from frameworks.onlyoffice.x2ttester.host_config import HostConfig

from .xml import *
from ...host_control import FileUtils


class X2tTesterXml(XML):
    def __init__(self, results_dir):
        self.results_dir = FileUtils.delete_last_slash(results_dir)
        self.host = HostConfig()
        self.core_dir = StaticData.core_dir()
        self.reports_dir = StaticData.reports_dir()
        self.fonts_dir = StaticData.fonts_dir()
        self.tmp_dir = StaticData.TMP_DIR
        self.documents_dir = FileUtils.delete_last_slash(StaticData.documents_dir())
        self.timeout = self.generate_timeout(settings.timeout)
        self.cores = settings.cores
        self.errors_only = self.parameter_validator(settings.errors_only)
        self.delete = self.parameter_validator(settings.delete)
        self.timestamp = self.parameter_validator(settings.timestamp)

    def create_tmp_files_list(self, files_array):
        return self.create_xml(self.files_list(files_array), xml_path=FileUtils.random_name(self.core_dir, 'xml'))

    def create_tmp_parameters(self, input_format, output_format, files_list_path, report_path):
        return self.create_xml(
            self.parameters(input_format, output_format, files_list_path, report_path),
            xml_path=FileUtils.random_name(self.core_dir, 'xml')
        )

    @staticmethod
    def files_list(files_array):
        root = ET.Element("files")
        for file in files_array:
            ET.SubElement(root, "file_name").text = file
        return root

    def parameters(self, input_format=None, output_format=None, files_list=None, report_path=None):
        root = ET.Element("Settings")
        ET.SubElement(root, "reportPath").text = report_path if report_path else self.reports_dir
        ET.SubElement(root, "inputDirectory").text = self.documents_dir
        ET.SubElement(root, "outputDirectory").text = self.results_dir
        ET.SubElement(root, "x2tPath").text = self.generate_x2t_path()
        ET.SubElement(root, "cores").text = self.generate_number_of_cores()
        ET.SubElement(root, "timeout").text = self.timeout
        if input_format:
            ET.SubElement(root, "input").text = input_format
        if output_format:
            ET.SubElement(root, "output").text = output_format
        if self.errors_only:
            ET.SubElement(root, "errorsOnly").text = self.errors_only
        if self.delete:
            ET.SubElement(root, "deleteOk").text = self.delete
        if self.timestamp:
            ET.SubElement(root, "timestamp").text = self.timestamp
        if files_list:
            ET.SubElement(root, "inputFilesList").text = files_list
        if exists(self.fonts_dir) and any(scandir(self.fonts_dir)):
            ET.SubElement(ET.SubElement(root, "fonts", system="0"), "directory").text = self.fonts_dir
        return root

    @staticmethod
    def parameter_validator(value):
        return value if value in ["1", "0"] else None

    def generate_x2t_path(self):
        x2t_path = join(self.core_dir, self.host.x2t)
        if not isfile(x2t_path):
            raise print(f'[bold red]Check x2t file_name, path: {x2t_path}[/]')
        return x2t_path

    def generate_number_of_cores(self):
        if self.cores == '':
            raise print('[bold red]The number of cores must be greater than 0')
        return self.cores

    @staticmethod
    def generate_timeout(value):
        return value if value else '0'
