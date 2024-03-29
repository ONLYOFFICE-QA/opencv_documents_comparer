# -*- coding: utf-8 -*-
from host_tools.utils import Shell
from rich import print

from frameworks.StaticData import StaticData
from frameworks.decorators import async_processing
from host_tools import HostInfo, File

from .handlers import PowerPointEvents


if HostInfo().os == 'windows':
    from win32com.client import Dispatch


def handler():
    PowerPointEvents().handler_for_thread()


@async_processing(target=handler)
class PresentationInfo:

    def __init__(self, file_path):
        self.tmp_dir = StaticData.tmp_dir_in_test
        self.tmp_file = File.make_tmp(file_path, self.tmp_dir)
        self.power_point = StaticData.powerpoint
        self.app = Dispatch("PowerPoint.application")
        self.presentation = self.__open(self.tmp_file)

    def __del__(self):
        self.presentation.Close() if self.presentation else ...
        self.app.Quit()
        Shell.call(f"taskkill /im {self.power_point}")
        File.delete(self.tmp_file, stdout=False)

    def slide_count(self) -> int:
        try:
            return len(self.presentation.Slides)
        except Exception as e:
            print(f"Exception when getting slide count: {e}")

    def __open(self, file_path: str):
        try:
            return self.app.Presentations.Open(file_path)
        except Exception as e:
            print(f"Exception when open file: {e}")
