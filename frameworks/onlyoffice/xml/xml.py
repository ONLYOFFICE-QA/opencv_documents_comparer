# -*- coding: utf-8 -*-

import xml.etree.cElementTree as ET


class XML:
    @staticmethod
    def create_xml(xml, xml_path=None):
        tree = ET.ElementTree(xml)
        ET.indent(tree, '  ')
        tree.write(xml_path, encoding="UTF-8", xml_declaration=True)
        return xml_path
