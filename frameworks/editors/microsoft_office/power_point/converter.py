# -*- coding: utf-8 -*-
from os.path import exists, join, dirname, basename, splitext

from frameworks.host_control import FileUtils, HostInfo

if HostInfo().os == 'windows':
    import win32com.client

class PowerPointConverter:

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        self.presentation = self.powerpoint.Presentations.Open(self.file_path)

    def __del__(self):
        self.presentation.Close()
        self.powerpoint.Quit()

    def to_png(self, output_dir: str = None) -> None:
        res_dir = output_dir if output_dir else self._generate_result_dir()
        FileUtils.create_dir(res_dir, silence=True)
        for i, slide in enumerate(self.presentation.Slides, start=1):
            slide.Export(join(res_dir, f"slide_{i}.png"), "PNG")

    def _generate_result_dir(self) -> str:
        extension = splitext(self.file_path)[1]
        dir_name =  basename(self.file_path).replace(extension, f"_{extension.replace('.', '')}")
        return join(dirname(self.file_path), dir_name)