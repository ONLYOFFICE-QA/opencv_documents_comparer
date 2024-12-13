# -*- coding: utf-8 -*-
from os.path import basename, isdir, splitext, join, dirname

from host_tools import File, HostInfo
from host_tools.utils import Dir
from rich.progress import track

from .x2ttester_test_config import X2ttesterTestConfig


class ResultsHandler:
    img_formats = ["png", "jpg", "bmp", "gif"]

    def __init__(self, test_config: X2ttesterTestConfig):
        self.config = test_config
        self.os = HostInfo().os

    def _get_paths(self, output_format: str) -> list:
        if output_format in self.img_formats:
            return Dir.get_paths(self.config.tmp_dir, end_dir=f".{output_format}")
        return File.get_paths(self.config.tmp_dir, extension=output_format)

    def run(self, result_path: bool | str = None, output_format: str = None) -> None:
        for output_format in output_format.strip().split(' ') if output_format else range(1):
            paths = self._get_paths(output_format if output_format else None)
            if paths:
                for file_path in track(paths, f"[cyan]|INFO| Copying {len(paths)} {output_format} files"):

                    _path_to = self._get_result_path(file_path, result_path, output_format)

                    Dir.create(_path_to, stdout=False) if not isdir(_path_to) else ...
                    name = basename(file_path)
                    File.copy(
                        file_path,
                        join(_path_to, name if len(name) < 200 else f'{name[:150]}.{splitext(name)[1]}'),
                        stdout=False
                    )

    def _get_input_format(self, file_path: str) -> str:
        if self.config.trough_conversion:
            return basename(dirname(dirname(file_path)))
        return dirname(file_path).split('.')[-1]

    def _get_result_path(self, file_path: str, result_path: str = None, output_format: str = None) -> str:

        if isinstance(result_path, str):
            return result_path

        input_format = self._get_input_format(file_path)
        return join(
            self.config.result_dir,
            f"{self.config.x2t_version}_"
            f"(dir_{input_format.lower().replace('.', '')}-{output_format})_"
            f"(os_{self.os})_"
            f"(mode_{'t-format' if self.config.trough_conversion else 'Default'})"
        )
