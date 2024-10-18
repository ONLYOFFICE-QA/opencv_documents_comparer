# -*- coding: utf-8 -*-
import re
from os import listdir
from os.path import basename, splitext, join, exists
from time import sleep
import numpy as np
from host_tools.utils import Process, Dir

from rich import print
from rich.progress import track

import config
from frameworks.StaticData import StaticData
from frameworks.editors import Document, PowerPoint, LibreOffice, Word, Excel
from frameworks.editors.onlyoffice import VersionHandler
from host_tools import File
from frameworks.image_handler import Image
from config import version
from frameworks.decorators import singleton, timer
from .tools import OpenerReport, CompareReport


@singleton
class CompareTest:
    def __init__(self):
        self.source_dir = config.source_docs
        self.source_screen_dir = StaticData.tmp_dir_source_img
        self.converted_screen_dir = StaticData.tmp_dir_converted_img
        self.tmp_dir = StaticData.tmp_dir_in_test
        self.result_folder = StaticData.result_dir()
        self.document_power_point = Document(PowerPoint())
        self.document_libre = Document(LibreOffice())
        self.document_word = Document(Word())
        self.document_excel = Document(Excel())
        self.image = Image()
        self.report_dir = join(StaticData.reports_dir(), VersionHandler(version).without_build, 'compare')
        self.opener_report = OpenerReport(join(self.report_dir, 'opener'))
        self.compare_report = CompareReport(self.report_dir)
        self.coefficient = 98
        self._clean_tmp_dirs()
        self._create_tmp_dirs()
        Process.terminate(StaticData.terminate_process)
        self.total, self.count = 0, 1

    @timer
    def run(self, converted_file_paths, source_ext, converted_ext):
        self.total = len(converted_file_paths)
        print(f"[bold green]\n{'-' * 90}\n|INFO| Compare on version: {version} is running.\n{'-' * 90}\n")
        for converted_file in converted_file_paths:
            source_file = File.get_paths(
                self.source_dir, names=[basename(converted_file).replace(f".{converted_ext}", f".{source_ext}")]
            )[0]
            print(f"[cyan]\n{'-' * 90}\n({self.count}/{self.total})[/] [green]In comparison test:[/] "
                  f"{basename(source_file)} [green]and[/] {basename(converted_file)}")
            converted_document_type = self._document_type(converted_file)
            page_amount = converted_document_type.page_amount(source_file if splitext(converted_file)[1] in self.document_libre.formats else converted_file)
            print(f"[bold blue] |INFO| Number of pages: {page_amount}")
            if not self.make_screen(converted_document_type, converted_file, self.converted_screen_dir, page_amount):
                Dir.delete(f'{self.converted_screen_dir}', clear_dir=True, stdout=False)
                continue
            self.make_screen(self._document_type(source_file), source_file, self.source_screen_dir, page_amount)
            self.compare_screens(converted_file, source_file)
            self.count += 1

    def make_screen(self, document_type: Document, path: str, screen_path: str, page_num: int | dict) -> bool | None:
        print(f'[bold green] |INFO| Getting Screenshots from: [cyan]{basename(path)}')
        tmp_file = File.make_tmp(path, self.tmp_dir)
        hwnd = document_type.open(tmp_file)
        if not isinstance(hwnd, int):
            self.opener_report.write(path, f"{hwnd}")
            document_type.close()
            document_type.delete(tmp_file)
            return False
        sleep(document_type.delay_after_open)
        errors_after_open = document_type.check_errors()
        if errors_after_open is True or errors_after_open is None:
            self.opener_report.write(path, 'ERROR')
            document_type.close(hwnd)
            document_type.delete(tmp_file)
            return False
        document_type.make_screenshots(hwnd, screen_path, page_num)
        document_type.close(hwnd)
        document_type.delete(tmp_file)
        return True

    def compare_screens(self, converted_file, source_file, coefficient: int = 98):
        for img_name in track(listdir(self.converted_screen_dir), description="[bold cyan]Comparing In Progress"):
            if exists(join(self.source_screen_dir, img_name)) and exists(join(self.converted_screen_dir, img_name)):
                sheet_num = img_name.split('_')[-1].replace('.png', '')
                similarity, difference, source_img, conv_img = self.find_difference(source_file, converted_file, img_name)
                similarity = round(similarity * 100, 3)
                msg = f" [cyan]|INFO| Compare screens. Sheet: {sheet_num} similarity[/]: {similarity} %"
                if similarity < coefficient:
                    self.image.put_text(source_img, f'Before sheet: {sheet_num}. Similarity: {similarity}%')
                    self.image.put_text(conv_img, f'After sheet: {sheet_num}. Similarity: {similarity}%')
                    self.save_result(source_file, converted_file, source_img, conv_img, img_name, difference)
                    print(f"{msg} [red] Aborted[/]")
                else:
                    print(f"{msg} [green] Passed[/]")
            else:
                print(f'Image {img_name} not found')
        self._clean_tmp_dirs()

    def find_difference(self, source_file, converted_file, img_name):
        source_img = self._find_sheet(self.image.read(join(self.source_screen_dir, img_name)), source_file)
        conv_img = self._find_sheet(self.image.read(join(self.converted_screen_dir, img_name)), converted_file)
        similarity, difference = self.image.find_difference(source_img, conv_img)

        if source_img.shape != conv_img.shape:
            source_img, conv_img = self.image.align_sizes(source_img, conv_img)

        return similarity, difference, source_img, conv_img

    def _find_sheet(self, img: np.ndarray, file_name) -> np.ndarray:
        if file_name.lower().endswith(self.document_excel.formats):
            return img

        try:
            processed_img = self.image.find_contours(img)
            if processed_img:
                return processed_img
            return img

        except Exception as ex:
            return img

    def save_result(self, source_file, converted_file, source_img, converted_img, img_name, difference):
        cnv_name, source_name = basename(converted_file), basename(source_file)
        path = join(self.report_dir, 'img_diff', re.sub(r"[\s\n\r.,\-=]+", '', cnv_name)[:35])
        Dir.create((join(path, 'gif'), join(path, 'screen')), stdout=False)
        File.copy(converted_file, join(path, cnv_name), stdout=False) if not exists(join(path, cnv_name)) else None
        File.copy(source_file, join(path, source_name), stdout=False) if not exists(join(path, source_name)) else None
        self.save_gif(path, img_name, source_img, converted_img, difference=difference)
        self.save_collage(path, img_name, source_img, converted_img)

    def save_gif(self, path, img_name, source_img, converted_img, difference=None):
        if not difference is None:
            source_img, converted_img = self.image.draw_differences(source_img, converted_img, difference)

        self.image.save_gif(
            join(path, 'gif', f"{splitext(img_name)[0]}_similarity.gif"),
            [source_img, converted_img],
            duration=1000
        )

    def save_collage(self, path, img_name, source_img, converted_img):
        self.image.save(join(path, 'screen', f"{img_name}_collage.png"), np.hstack([source_img, converted_img]))

    def _create_tmp_dirs(self):
        Dir.create((self.source_screen_dir, self.converted_screen_dir), stdout=False)

    def _clean_tmp_dirs(self):
        Dir.delete(
            (self.source_screen_dir, self.converted_screen_dir, self.tmp_dir),
            clear_dir=True,
            stdout=False
        )

    def _document_type(self, file_path):
        if file_path.lower().endswith(self.document_excel.formats + ('.ods',)):
            return self.document_excel
        elif file_path.lower().endswith(self.document_word.formats + ('.odt',)):
            return self.document_word
        elif file_path.lower().endswith(self.document_libre.formats):
            return self.document_libre
        elif file_path.lower().endswith(self.document_power_point.formats):
            return self.document_power_point
        else:
            raise print(f"[red]|WARNING| Editor not found to open the extension: {splitext(file_path)[1]}")

    @staticmethod
    def getting_formats(direction: str | None = None) -> tuple:
        if direction:
            if '-' in direction:
                return direction.split('-')[0], direction.split('-')[1]
            return None, direction
        return None, None
