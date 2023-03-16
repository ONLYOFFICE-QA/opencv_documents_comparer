# -*- coding: utf-8 -*-
from os.path import join
from rich import print

from .xml import *
from ...StaticData import StaticData


class X2tLibsXML(XML):

    def __init__(self, x2t_dir: str = StaticData.core_dir()):
        self.x2t_dir = x2t_dir

    @staticmethod
    def generate_doc_renderer_config():
        settings = ET.Element("Settings")
        ET.SubElement(settings, "file_name").text = './sdkjs/common/Native/native.js'
        ET.SubElement(settings, "file_name").text = './sdkjs/common/Native/jquery_native.js'
        ET.SubElement(settings, "allfonts").text = './fonts/AllFonts.js'
        ET.SubElement(settings, "file_name").text = './sdkjs/vendor/xregexp/xregexp-all-min.js'
        ET.SubElement(settings, "sdkjs").text = './sdkjs'
        return settings

    def create_doc_renderer_config(self):
        self.create_xml(self.generate_doc_renderer_config(), xml_path=join(self.x2t_dir, 'DoctRenderer.config'))
        print("[green]|INFO| DocRenderer.confing created")
