# -*- coding: utf-8 -*-
import xml.etree.cElementTree as ET
from os import scandir
from os.path import join, isfile, exists

from rich import print

import settings as config
from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions
from framework.singleton import singleton


@singleton
class XmlActions:
    def __init__(self):
        self.host = HostActions()

    def generate_x2t_path(self):
        if not isfile(join(ProjectConfig.core_dir(), self.host.x2t)):
            raise print(f'[bold red]Check x2t File, path: {ProjectConfig.core_dir()}/{self.host.x2t}[/]')
        return join(ProjectConfig.core_dir(), self.host.x2t)

    @staticmethod
    def generate_number_of_cores():
        if config.cores == '':
            raise print('[bold red]Please enter the number of cores in settings.py')
        return config.cores

    def generate_doc_renderer_config(self):
        settings = ET.Element("Settings")
        ET.SubElement(settings, "file").text = './sdkjs/common/Native/native.js'
        ET.SubElement(settings, "file").text = './sdkjs/common/Native/jquery_native.js'
        ET.SubElement(settings, "allfonts").text = './fonts/AllFonts.js'
        ET.SubElement(settings, "file").text = './sdkjs/vendor/xregexp/xregexp-all-min.js'
        ET.SubElement(settings, "sdkjs").text = './sdkjs'
        self.write_to_xml(settings, join(ProjectConfig.core_dir(), 'DoctRenderer.config'))

    def generate_files_list(self):
        root = ET.Element("files")
        for file in config.files_array:
            ET.SubElement(root, "file").text = file
        return self.write_to_xml(root)

    @staticmethod
    def generate_input_dir():
        return FileUtils.delete_last_slash(ProjectConfig.documents_dir())

    @staticmethod
    def generate_output_dir():
        return FileUtils.delete_last_slash(ProjectConfig.tmp_result_dir())

    def generate_report_name(self):
        report_name = f"{x2tversion}_{input_format}_{output_format}.csv"
        pass

    def generate_report_path(self, input_format, output_format, x2tversion):
        ProjectConfig.CSTM_REPORT_DIR = join(ProjectConfig.reports_dir(), self.host.os, f"conversion")
        FileUtils.create_dir(ProjectConfig.CSTM_REPORT_DIR)
        report_name = f"{x2tversion}_{input_format}_{output_format}.csv"
        return join(ProjectConfig.CSTM_REPORT_DIR, report_name)

    def generate_x2ttester_parameters(self, input_format=None, output_format=None, files_list_path=None, x2tversion=''):
        root = ET.Element("Settings")
        ET.SubElement(root, "reportPath").text = self.generate_report_path(input_format, output_format, x2tversion)
        ET.SubElement(root, "inputDirectory").text = self.generate_input_dir()
        ET.SubElement(root, "outputDirectory").text = self.generate_output_dir()
        ET.SubElement(root, "x2tPath").text = self.generate_x2t_path()
        ET.SubElement(root, "cores").text = self.generate_number_of_cores()
        if input_format:
            ET.SubElement(root, "input").text = input_format
        if output_format:
            ET.SubElement(root, "output").text = output_format
        if config.errors_only in ["1", "0"]:
            ET.SubElement(root, "errorsOnly").text = config.errors_only
        if config.delete in ["1", "0"]:
            ET.SubElement(root, "deleteOk").text = config.delete
        if config.timestamp in ["1", "0"]:
            ET.SubElement(root, "timestamp").text = config.timestamp
        if files_list_path:
            ET.SubElement(root, "inputFilesList").text = files_list_path
        if exists(ProjectConfig.fonts_dir()) and any(scandir(ProjectConfig.fonts_dir())):
            fonts = ET.SubElement(root, "fonts", system="0")
            ET.SubElement(fonts, "directory").text = ProjectConfig.fonts_dir()
        return self.write_to_xml(root)

    @staticmethod
    def write_to_xml(xml, path_to_xml=None):
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        path = FileUtils.random_name(ProjectConfig.core_dir(), 'xml') if not path_to_xml else path_to_xml
        tree.write(path, encoding="UTF-8", xml_declaration=True)
        return path
