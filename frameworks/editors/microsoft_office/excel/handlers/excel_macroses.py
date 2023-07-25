# -*- coding: utf-8 -*-
from os.path import join
from time import sleep

from frameworks.StaticData import StaticData
from frameworks.editors.microsoft_office.excel.handlers import ExcelEvents
from frameworks.host_control import KeyActions


class ExcelMacroses:

    @staticmethod
    def prepare_for_test():
        if KeyActions.click(join(StaticData.project_dir, 'assets', 'image_templates', 'excel', 'turn_on_content.png')):
            sleep(1)
            ExcelEvents().when_turn_on_content()
