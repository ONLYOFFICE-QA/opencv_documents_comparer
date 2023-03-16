# -*- coding: utf-8 -*-
from rich import print

from frameworks.StaticData import StaticData
from frameworks.image_handler import CompareImage
from frameworks.libre_office import LibreOffice
from frameworks.microsoft_office import PowerPoint


class OdpPptxCompare(PowerPoint, LibreOffice):
    def run_compare(self, list_of_files):
        self.doc_helper.terminate_process()
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.endswith((".pptx", ".PPTX")):
                continue
            self.doc_helper.preparing_files_for_compare_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            if not self.get_slide_count(self.doc_helper.tmp_file_for_get_statistic):
                continue
            self.open_presentation_with_cmd(self.doc_helper.tmp_converted_file)
            if not self.errors_handler_when_opening():
                self.close_presentation_with_hotkey()
                continue
            self.get_screenshot(StaticData.TMP_DIR_CONVERTED_IMG)
            self.close_presentation_with_hotkey()

            print(f'[bold green]In test[/] {self.doc_helper.source_file}')
            self.open_libre_office_with_cmd(self.doc_helper.tmp_source_file)
            self.get_screenshot_odp(StaticData.TMP_DIR_SOURCE_IMG, self.slide_count)
            self.close_libre()

            CompareImage()
            self.doc_helper.tmp_cleaner()
