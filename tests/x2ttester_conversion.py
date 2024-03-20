# -*- coding: utf-8 -*-
from os.path import join, dirname, realpath
from rich import print

import config
from frameworks.StaticData import StaticData
from frameworks.decorators import timer
from frameworks.editors import X2tTester, X2tTesterData
from frameworks.editors.onlyoffice import VersionHandler, X2t
from host_tools import File, Dir
from .tools import X2ttesterReport
from .tools.results_handler import ResultsHandler


class X2tTesterConversion:
    QUICK_CHECK_FILES_PATH = join(dirname(realpath(__file__)), 'assets', 'quick_check_files.json')
    EXTENSIONS_FILE_PATH = join(dirname(realpath(__file__)), 'assets', 'extension_array.json')

    def __init__(
            self,
            direction: str | None = None,
            version: str = None,
            trough_conversion: bool = False,
            env_off: bool = False
    ):
        self.env_off = env_off
        self.trough_conversion = trough_conversion
        self.input_formats, self.output_formats = self._getting_formats(direction)
        self.tmp_dir = join(StaticData.tmp_dir, 'cnv')
        self.x2t_dir = StaticData.core_dir()
        self.report = X2ttesterReport()
        self.x2t_version = VersionHandler(version if version else X2t.version(StaticData.core_dir())).version
        self.report_path = self.report.tmp_file()
        self.result_dir = StaticData.result_dir()
        self.x2ttester = X2tTester(self._get_x2ttester_data())
        self.results_handler = self._get_results_handler()
        Dir.delete(StaticData.tmp_dir, clear_dir=True, stdout=False, stderr=False)

    def run(self, results_path: bool | str = False, list_xml: str = None) -> str:
        """
        :param results_path: bool | str - takes True, False, String
        if String then it is path to folder where files will be copied
        if True the path will be generated automatically, based on x2t version and configurable options
        :param list_xml: str — the path to the xml file with the names of the files for the test
        """
        self.x2ttester.conversion(self.input_formats, self.output_formats, listxml_path=list_xml)
        self.results_handler.run(results_path) if results_path is not False else ...
        return File.last_modified(dirname(self.report_path))

    @timer
    def from_extension_json(self) -> str | None:
        reports = []
        for output_format, inout_formats in File.read_json(self.EXTENSIONS_FILE_PATH).items():
            self.output_formats = output_format
            self.input_formats = " ".join(inout_formats) if inout_formats else None

            print(
                f"[green]|INFO| Conversion direction: "
                f"[cyan]{self.input_formats if self.input_formats else 'All'} "
                f"[red]to [cyan]{self.output_formats}"
            )

            reports.append(self.run(results_path=True))
            Dir.delete(self.tmp_dir, clear_dir=True)

        return self.report.merge_reports(reports, self.x2t_version)

    def get_quick_check_files(self) -> list:
        return sum([array for _, array in File.read_json(self.QUICK_CHECK_FILES_PATH).items()], [])

    @timer
    def from_files_list(self, files: list) -> str:
        xml = self.x2ttester.xml.create(
            self.x2ttester.xml.files_list(files),
            File.unique_name(self.x2t_dir, 'xml')
        )
        report = self.run(list_xml=xml)
        File.delete(xml)
        return report

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
            report_path=self.report_path,
            timeout=config.timeout,
            timestamp=config.timestamp,
            delete=config.delete,
            errors_only=config.errors_only,
            trough_conversion=self.trough_conversion,
            environment_off=self.env_off
        )

    def _get_results_handler(self) -> ResultsHandler:
        return ResultsHandler(
            self.output_formats,
            self.tmp_dir,
            self.result_dir,
            self.x2t_version,
            self.trough_conversion
        )
