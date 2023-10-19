# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import basename, dirname, join
from time import sleep

from rich import print

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton, timer
from frameworks.editors import Document, PowerPoint, LibreOffice, Word, Excel
from frameworks.editors.onlyoffice import VersionHandler
from host_control import File, Dir, Window, Process
from .tools import OpenerReport


@singleton
class OpenTests:
    def __init__(self, version: str, continue_test: bool = True):
        self.version = VersionHandler(version)
        self.continue_test: bool = continue_test
        self.tmp_dir = StaticData.tmp_dir_in_test
        self.document_power_point = Document(PowerPoint())
        self.document_libre = Document(LibreOffice())
        self.document_word = Document(Word())
        self.document_excel = Document(Excel())
        self.report = OpenerReport(self._generate_report_path())
        self.total, self.count = 0, 1
        self._prepare_test()

    @timer
    def run(self, file_paths: list, tg_msg: bool | str = False) -> None:
        testing_paths = self._generate_testing_paths(file_paths)
        self.total = len(testing_paths)
        print(f'[bold green]\n{"-" * 90}\n|INFO| Opener on version: {self.version.version} is running.\n{"-" * 90}\n')
        for file_path in testing_paths:
            if file_path.lower().endswith(self.document_excel.formats):
                self.opening_test(self.document_excel, file_path)
            elif file_path.lower().endswith(self.document_word.formats):
                self.opening_test(self.document_word, file_path)
            elif file_path.lower().endswith(self.document_libre.formats):
                self.opening_test(self.document_libre, file_path)
            elif file_path.lower().endswith(self.document_power_point.formats):
                self.opening_test(self.document_power_point, file_path)
            else:
                print(f"[red]|WARNING| Editor not found to open the file: {file_path}")
            self.count += 1
        self.report.handler(tg_msg)

    def opening_test(self, document_type: Document, file_path: str) -> bool | None:
        print(
            f'[cyan]({self.count}/{self.total})[/] [green]In opening test:[/] '
            f'[cyan]{basename(dirname(file_path))}[/][red]/[/]{basename(file_path)}'
        )

        tmp_file = File.make_tmp(file_path, File.unique_name(self.tmp_dir))
        hwnd = document_type.open(tmp_file)

        if not isinstance(hwnd, int):
            self.report.write(file_path, f"{hwnd}")
            document_type.close()
            return document_type.delete(tmp_file)

        Window.set_size(hwnd, 0, 0, 900, 800) if self.count == 1 else ...
        sleep(document_type.delay_after_open)

        errors_after_open = document_type.check_errors()
        if errors_after_open is True or errors_after_open is None:
            self.report.write(file_path, 'ERROR')
            document_type.close(hwnd)
            return document_type.delete(tmp_file)

        document_type.close(hwnd)
        if document_type.delete(dirname(tmp_file)) is False:
            return self.report.write(file_path, 'CANT_DELETE')

        self.report.write(file_path, 0)

    def getting_formats(self, direction: str | None = None) -> tuple:
        if direction:
            if '-' in direction:
                return direction.split('-')[0], direction.split('-')[1]
            elif 'libre' in direction:
                print(f"[bold green]|INFO| Filter by [red]LibreOffice[/][bold green] formats")
                return None, self.document_libre.formats
            elif 'powerpoint' in direction or direction == 'pp':
                print(f"[bold green]|INFO| Filter by [red]PowerPoint[/][bold green] formats")
                return None, self.document_power_point.formats
            elif 'word' in direction:
                print(f"[bold green]|INFO| Filter by [red]Word[/][bold green] formats")
                return None, self.document_word.formats
            elif 'excel' in direction:
                print(f"[bold green]|INFO| Filter by [red]Excel[/][bold green] formats")
                return None, self.document_excel.formats
            elif 'msoffice' in direction:
                formats = self.document_power_point.formats + self.document_word.formats + self.document_excel.formats
                print(f"[bold green]|INFO| Filter by [red]MS Office[/][bold green] formats")
                return None, formats
            else:
                return None, direction
        return None, None

    def _generate_report_path(self):
        report_dir = join(StaticData.reports_dir(), self.version.without_build, 'opener')
        if self.continue_test is True:
            return join(report_dir, f"{self.version.version}")
        return join(
            report_dir,
            "tmp_reports",
            f"{self.version.version}_opener_{datetime.now().strftime('%H_%M_%S')}.csv"
        )

    def _generate_testing_paths(self, file_paths: list) -> list:
        if self.continue_test is True:
            tested_files = self.report.tested_files()
            return [path for path in file_paths if join(basename(dirname(path)), basename(path)) not in tested_files]
        return file_paths

    def _prepare_test(self):
        Dir.delete(self.tmp_dir, clear_dir=True, stdout=False, stderr=False)
        Dir.create(self.tmp_dir, stdout=False)
        Process.terminate(StaticData.terminate_process)
