# -*- coding: utf-8 -*-
import xml.etree.cElementTree as ET
from os import scandir
from os.path import join, isfile, exists
from re import sub

from rich import print

import settings
from framework.StaticData import StaticData
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions


class XmlActions:
    def __init__(self):
        self.host = HostActions()

    def generate_x2t_path(self):
        if not isfile(join(StaticData.core_dir(), self.host.x2t)):
            raise print(f'[bold red]Check x2t File, path: {StaticData.core_dir()}/{self.host.x2t}[/]')
        return join(StaticData.core_dir(), self.host.x2t)

    @staticmethod
    def generate_number_of_cores():
        if settings.cores == '':
            raise print('[bold red]Please enter the number of cores in settings.py')
        return settings.cores

    def generate_doc_renderer_config(self):
        settings = ET.Element("Settings")
        ET.SubElement(settings, "file").text = './sdkjs/common/Native/native.js'
        ET.SubElement(settings, "file").text = './sdkjs/common/Native/jquery_native.js'
        ET.SubElement(settings, "allfonts").text = './fonts/AllFonts.js'
        ET.SubElement(settings, "file").text = './sdkjs/vendor/xregexp/xregexp-all-min.js'
        ET.SubElement(settings, "sdkjs").text = './sdkjs'
        self.write_to_xml(settings, join(StaticData.core_dir(), 'DoctRenderer.config'))

    def generate_files_list(self):
        root = ET.Element("files")
        for file in settings.files_array:
            ET.SubElement(root, "file").text = file
        return self.write_to_xml(root)

    @staticmethod
    def generate_input_dir():
        return FileUtils.delete_last_slash(StaticData.documents_dir())

    @staticmethod
    def generate_output_dir():
        return FileUtils.delete_last_slash(StaticData.tmp_result_dir())

    @staticmethod
    def generate_major_version(full_version):
        if len([i for i in full_version.split('.') if i]) == 4:
            return sub(r'(\d+).(\d+).(\d+).(\d+)', r'\1.\2.\3', full_version)
        raise print("[bold red]|WARNING| The version is entered incorrectly")

    def generate_report_dir(self, x2t_version):
        major_version = self.generate_major_version(x2t_version)
        reports_dir = join(StaticData.reports_dir(), major_version, self.host.os, f"conversion")
        FileUtils.create_dir(reports_dir)
        return reports_dir

    # @param [String] input_format
    # @param [String] output_format
    # @param [String] x2t_version
    # @return [String] path to report.csv and tmp report directory.
    @staticmethod
    def generate_report_paths(input_format, output_format, x2t_version):
        tmp_report_dir = FileUtils.random_name(StaticData.TMP_DIR)
        report_name = f"{x2t_version}_{input_format}_{output_format}.csv"
        FileUtils.create_dir(tmp_report_dir, silence=True)
        return join(tmp_report_dir, report_name), tmp_report_dir

    @staticmethod
    def generate_timeout():
        return settings.timeout if settings.timeout else '0'

    def generate_x2ttester_parameters(self, input_format=None, output_format=None, files_list=None, report_path=''):
        root = ET.Element("Settings")
        ET.SubElement(root, "reportPath").text = report_path if report_path else StaticData.reports_dir()
        ET.SubElement(root, "inputDirectory").text = self.generate_input_dir()
        ET.SubElement(root, "outputDirectory").text = self.generate_output_dir()
        ET.SubElement(root, "x2tPath").text = self.generate_x2t_path()
        ET.SubElement(root, "cores").text = self.generate_number_of_cores()
        ET.SubElement(root, "timeout").text = self.generate_timeout()
        if input_format:
            ET.SubElement(root, "input").text = input_format
        if output_format:
            ET.SubElement(root, "output").text = output_format
        if settings.errors_only in ["1", "0"]:
            ET.SubElement(root, "errorsOnly").text = settings.errors_only
        if settings.delete in ["1", "0"]:
            ET.SubElement(root, "deleteOk").text = settings.delete
        if settings.timestamp in ["1", "0"]:
            ET.SubElement(root, "timestamp").text = settings.timestamp
        if files_list:
            ET.SubElement(root, "inputFilesList").text = files_list
        if exists(StaticData.fonts_dir()) and any(scandir(StaticData.fonts_dir())):
            fonts = ET.SubElement(root, "fonts", system="0")
            ET.SubElement(fonts, "directory").text = StaticData.fonts_dir()
        return self.write_to_xml(root)

    @staticmethod
    def write_to_xml(xml, path_to_xml=None):
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        path = FileUtils.random_name(StaticData.core_dir(), 'xml') if not path_to_xml else path_to_xml
        tree.write(path, encoding="UTF-8", xml_declaration=True)
        return path
