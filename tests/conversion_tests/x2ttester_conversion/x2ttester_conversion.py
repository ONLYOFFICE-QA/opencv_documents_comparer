# -*- coding: utf-8 -*-
import os
from os.path import join, dirname
from rich import print


from frameworks.decorators import timer
from frameworks.editors import X2tTester, X2tTesterConfig
from host_tools import File, Dir

from .x2ttester_report import X2ttesterReport
from .results_handler import ResultsHandler
from .x2ttester_test_config import X2ttesterTestConfig



class X2tTesterConversion:
    TEST_ASSETS_DIR = join(os.getcwd(), 'tests', 'assets')
    QUICK_CHECK_FILES_PATH = join(TEST_ASSETS_DIR, 'quick_check_files.json')
    EXTENSIONS_FILE_PATH = join(TEST_ASSETS_DIR, 'extension_array.json')

    def __init__(self, test_config: X2ttesterTestConfig):
        self.config = test_config
        self.input_formats, self.output_formats = self._getting_formats(self.config.direction)
        self.report = X2ttesterReport(self.config)
        self.x2ttester = X2tTester(config=X2tTesterConfig(**self.config.model_dump()))
        self.results_handler = ResultsHandler(self.config)

    def run(self, results_path: bool | str = False, list_xml: str = None) -> str:
        """
        :param results_path: bool | str - takes True, False, String
        if String then it is path to folder where files will be copied
        if True the path will be generated automatically, based on x2t version and configurable options
        :param list_xml: str — the path to the xml file with the names of the files for the test
        """
        self.x2ttester.conversion(self.input_formats, self.output_formats, listxml_path=list_xml)
        self.results_handler.run(results_path, self.output_formats) if results_path is not False else ...
        return File.last_modified(dirname(self.config.report_path))

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
            Dir.delete(self.config.tmp_dir, clear_dir=True)

        return self.report.merge_reports(reports, self.config.x2t_version)

    def get_quick_check_files(self) -> list:
        return sum([array for _, array in File.read_json(self.QUICK_CHECK_FILES_PATH).items()], [])

    @timer
    def from_files_list(self, files: list) -> str:
        xml = self.x2ttester.xml.create(
            self.x2ttester.xml.files_list(files),
            File.unique_name(self.config.core_dir, 'xml')
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
