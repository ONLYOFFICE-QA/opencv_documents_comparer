# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath
from rich import print

import xml.etree.cElementTree as ET
from frameworks.StaticData import StaticData
from host_control import FileUtils


class X2tLibsXML:
    def __init__(self, x2t_dir: str = StaticData.core_dir()):
        self.x2t_dir = x2t_dir
        self.x2t_lib_config = FileUtils.read_json(join(dirname(realpath(__file__)), 'assets', 'x2t_libs_paths.json'))

    def generate_doc_renderer_config(self):
        root = ET.Element("Settings")
        ET.SubElement(root, "file").text = self.x2t_lib_config['native']
        ET.SubElement(root, "file").text = self.x2t_lib_config['jquery_native']
        ET.SubElement(root, "allfonts").text = self.x2t_lib_config['AllFonts']
        ET.SubElement(root, "file").text = self.x2t_lib_config['xregexp_all_min']
        ET.SubElement(root, "sdkjs").text = self.x2t_lib_config['sdkjs']
        return root

    def create_doc_renderer_config(self):
        self._create(self.generate_doc_renderer_config(), xml_path=join(self.x2t_dir, 'DoctRenderer.config'))
        print("[green]|INFO| DocRenderer.confing created")

    @staticmethod
    def _create(xml: ET.Element, xml_path: str) -> str:
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        tree.write(xml_path, encoding="UTF-8", xml_declaration=True)
        return xml_path
