# -*- coding: utf-8 -*-
from os.path import join, basename, dirname, splitext, abspath, isdir

from rich.progress import track
from rich import print

import config
from frameworks.StaticData import StaticData
from frameworks.decorators import timer
from frameworks.editors import X2tTester, X2tTesterData
from frameworks.editors.onlyoffice import VersionHandler, X2t
from host_control import File, Dir
from .tools import X2ttesterReport


class X2tTesterConversion:
    def __init__(self, direction: str | None = None, version: str = None, trough_conversion: bool = False):
        self.trough_conversion = trough_conversion
        self.input_formats, self.output_formats = self._getting_formats(direction)
        self.extensions = File.read_json(f"{dirname(abspath(__file__))}/assets/extension_array.json")
        self.tmp_dir = join(StaticData.tmp_dir, 'cnv')
        self.result_dir = StaticData.result_dir()
        self.x2t_dir = StaticData.core_dir()
        self.report = X2ttesterReport()
        self.x2t_version = VersionHandler(version if version else X2t.version(StaticData.core_dir())).version
        self.img_formats = ["png", "jpg", "bmp", "gif"]
        self.data = self._get_x2ttester_data()
        self.x2ttester = X2tTester(self.data)
        Dir.delete(self.tmp_dir, clear_dir=True, stdout=False, stderr=False)


    def run(self, results_path: bool | str = False, list_xml: str = None) -> str:
        """
        :param results_path: bool | str - takes True, False, String
        if String then it is path to folder where files will be copied
        if True the path will be generated automatically, based on x2t version and configurable options
        :param list_xml: str — the path to the xml file with the names of the files for the test
        """
        self.x2ttester.conversion(self.input_formats, self.output_formats, listxml_path=list_xml)
        self._handle_results(results_path) if results_path is not False else ...
        return File.last_modified(dirname(self.data.report_path))

    @timer
    def from_extension_json(self) -> str | None:
        reports = []
        for extensions in self.extensions.items():
            self.output_formats = extensions[0]
            self.input_formats = " ".join(extensions[1]) if extensions[1] else None
            print(f"[green]|INFO| Conversion direction:"
                  f" [cyan]{self.input_formats if self.input_formats else 'All'} [red]to [cyan]{self.output_formats}")
            reports.append(self.run(results_path=True))
            Dir.delete(self.tmp_dir, clear_dir=True)
        return self.report.merge_reports(reports, self.x2t_version)

    @timer
    def from_files_list(self, files: list) -> str:
        xml = self.x2ttester.xml.create(
            self.x2ttester.xml.files_list(files),
            File.unique_name(self.x2t_dir, 'xml')
        )
        report = self.run(list_xml=xml)
        File.delete(xml)
        return report

    def _get_paths(self, output_format: str) -> list:
        if output_format in self.img_formats:
            return Dir.get_paths(self.tmp_dir, end_dir=f".{output_format}")
        return File.get_paths(self.tmp_dir, extension=output_format)

    def _handle_results(self, result_path: str = None) -> None:
        for output_formats in self.output_formats.strip().split(' ') if self.output_formats else range(1):
            paths = self._get_paths(output_formats if self.output_formats else None)
            if paths:
                for file_path in track(paths, f"[cyan]|INFO| Copying {len(paths)} {output_formats} files"):
                    _path_to = self._result_path(result_path, splitext(dirname(file_path))[1])
                    name = basename(file_path)
                    if splitext(name)[1].lower().replace('.', '') in self.img_formats:
                        _path_to = join(_path_to, name if len(name) < 200 else f'{name[:100]}.{splitext(name)[1]}')
                    Dir.create(_path_to, stdout=False)
                    File.copy(
                        file_path,
                        join(_path_to, name if len(name) < 200 else f'{name[:150]}.{splitext(name)[1]}'),
                        stdout=False
                    )

    def _result_path(self, result_path: str, input_format: str = None) -> str:
        if isinstance(result_path, str):
            return result_path
        return join(self.result_dir,
                    f"{self.x2t_version}_{input_format.lower().replace('.', '')}_{self.output_formats}")

    @staticmethod
    def _getting_formats(direction: str | None = None) -> tuple[None | str, None | str]:
        if direction:
            if '-' in direction:
                return direction.split('-')[0], direction.split('-')[1]
            return None, direction
        return None, None

    def _get_x2ttester_data(self):
        return X2tTesterData(
            cores=config.cores,
            input_dir=StaticData.documents_dir(),
            output_dir=self.tmp_dir,
            x2ttester_dir=self.x2t_dir,
            fonts_dir=StaticData.fonts_dir(),
            report_path=self.report.tmp_file(),
            timeout=config.timeout,
            timestamp=config.timestamp,
            delete=config.delete,
            errors_only=config.errors_only,
            trough_conversion=self.trough_conversion
        )
