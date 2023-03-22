# -*- coding: utf-8 -*-

import subprocess as sb
from os.path import join

import settings
from frameworks.StaticData import StaticData


class LibreOffice:
    def __init__(self):
        self.libre_path = join(settings.libre_office, StaticData.LIBRE)

    def open(self, file_path):
        sb.Popen(f"{self.libre_path} -o {file_path}")
