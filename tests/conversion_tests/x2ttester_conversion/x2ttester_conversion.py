# -*- coding: utf-8 -*-
from os.path import dirname
from pathlib import Path

from rich import print


from frameworks.decorators import timer
from frameworks.editors import X2tTester, X2tTesterConfig
from host_tools import File, Dir

from .x2ttester_report import X2ttesterReport
from .results_handler import ResultsHandler
from .x2ttester_test_config import X2ttesterTestConfig
from .conversion_test_tinfo import ConversionTestInfo


class X2tTesterConversion:
    TEST_ASSETS_DIR = Path(__file__).resolve().parents[2] / 'assets'
    QUICK_CHECK_FILES_PATH = str(TEST_ASSETS_DIR / 'quick_check_files.json')
    EXTENSIONS_FILE_PATH = str(TEST_ASSETS_DIR / 'extension_array.json')
    EXCEPTIONS_JSON = str(TEST_ASSETS_DIR / 'conversion_exception.json')

    def __init__(self, test_config: X2ttesterTestConfig):
        self.config = test_config
        self.report = X2ttesterReport(self.config, exceptions_json=File.read_json(self.EXCEPTIONS_JSON))
        self.x2ttester = X2tTester(config=X2tTesterConfig(**self.config.model_dump()))
        self.results_handler = ResultsHandler(self.config)
        self.info = ConversionTestInfo(test_config=self.config)

    def run(self, results_path: bool | str = False, list_xml: str = None) -> str:
        """
        :param results_path: bool | str - takes True, False, String
        if String then it is path to folder where files will be copied
        if True the path will be generated automatically, based on x2t version and configurable options
        :param list_xml: str — the path to the xml file with the names of the files for the test
        """
        self.x2ttester.conversion(self.config.input_formats, self.config.output_formats, listxml_path=list_xml)
        self.results_handler.run(results_path, self.config.output_formats) if results_path is not False else ...
        return File.last_modified(dirname(self.config.report_path))

    @timer
    def from_extension_json(self) -> str | None:
        reports = []
        for output_format, inout_formats in File.read_json(self.EXTENSIONS_FILE_PATH).items():
            self.config.output_formats = output_format
            self.config.input_formats = " ".join(inout_formats) if inout_formats else None

            print(
                f"[green]|INFO| Conversion direction: "
                f"[cyan]{self.config.input_formats if self.config.input_formats else 'All'} "
                f"[red]to [cyan]{self.config.output_formats}"
            )

            reports.append(self.run(results_path=True))
            self.x2ttester.xml.config.report_path = self.config.get_tmp_report_path()
            Dir.delete(self.config.output_dir, clear_dir=True)

        return self.report.merge_reports(reports)

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
