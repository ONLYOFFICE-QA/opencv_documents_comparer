# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath
from rich import print

import xml.etree.cElementTree as ET
from frameworks.StaticData import StaticData
from host_tools import File


class X2tLibsXML:
    """
    Class for generating XML configuration files for X2tLibs.

    This class provides methods for generating XML configuration files required for X2tLibs.

    Attributes:
        x2t_dir (str): Path to the X2t directory.
        x2t_lib_config (dict): Configuration data for X2tLibs paths.
    """

    def __init__(self, x2t_dir: str = StaticData.core_dir()):
        """
        Initializes the X2tLibsXML object.
        :param x2t_dir: Path to the X2t directory. Defaults to StaticData.core_dir().
        """
        self.x2t_dir = x2t_dir
        self.x2t_lib_config = File.read_json(join(dirname(realpath(__file__)), 'assets', 'x2t_libs_paths.json'))

    def generate_doc_renderer_config(self):
        """
        Generates the document renderer configuration XML.
        :return: XML element containing document renderer configuration.
        """
        root = ET.Element("Settings")
        ET.SubElement(root, "file").text = self.x2t_lib_config['native']
        ET.SubElement(root, "file").text = self.x2t_lib_config['jquery_native']
        ET.SubElement(root, "allfonts").text = self.x2t_lib_config['AllFonts']
        ET.SubElement(root, "file").text = self.x2t_lib_config['xregexp_all_min']
        ET.SubElement(root, "sdkjs").text = self.x2t_lib_config['sdkjs']
        return root

    def create_doc_renderer_config(self):
        """
        Creates the document renderer configuration XML file.
        :return: Path to the created XML file.
        """
        self._create(self.generate_doc_renderer_config(), xml_path=join(self.x2t_dir, 'DoctRenderer.config'))
        print("[green]|INFO| DocRenderer.confing created")

    @staticmethod
    def _create(xml: ET.Element, xml_path: str) -> str:
        """
        Creates an XML file with the provided XML tree.
        :param xml: XML tree to write into the file.
        :param xml_path: Path to save the XML file.
        :return: Path to the created XML file.
        """
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        tree.write(xml_path, encoding="UTF-8", xml_declaration=True)
        return xml_path
