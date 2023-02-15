# -*- coding: utf-8 -*-
from os.path import join

import pyautogui as pg

from configurations.project_configurator import ProjectConfig


class KeyActions:

    @staticmethod
    def click(path_to_image):
        try:
            pg.click(join(ProjectConfig.PROJECT_DIR, 'data', 'image_templates', path_to_image))
            return True
        except TypeError:
            return False
