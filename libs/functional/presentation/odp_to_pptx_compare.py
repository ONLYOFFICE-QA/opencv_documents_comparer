# -*- coding: utf-8 -*-
from data.project_configurator import ProjectConfig
from framework.libre_office import LibreOffice
from framework.power_point import PowerPoint
from framework.compare_image import CompareImage
from rich import print


class OdpPptxCompare(PowerPoint, LibreOffice):
    def run_compare(self, list_of_files):
        for self.doc_helper.converted_file in list_of_files:
            if not self.doc_helper.converted_file.endswith((".pptx", ".PPTX")):
                continue
            self.doc_helper.preparing_files_for_compare_test()
            print(f'[bold green]In test[/] {self.doc_helper.converted_file}')
            if not self.get_slide_count():
                continue
            self.open_presentation_with_cmd(self.doc_helper.tmp_converted_file)
            if not self.errors_handler_when_opening():
                self.close_presentation_with_hotkey()
                continue
            self.get_screenshot(ProjectConfig.TMP_DIR_CONVERTED_IMG)
            self.close_presentation_with_hotkey()

            print(f'[bold green]In test[/] {self.doc_helper.source_file}')
            self.open_libre_office_with_cmd(self.doc_helper.tmp_source_file)
            self.get_screenshot_odp(ProjectConfig.TMP_DIR_SOURCE_IMG, self.slide_count)
            self.close_libre()

            CompareImage()
            self.doc_helper.tmp_cleaner()
