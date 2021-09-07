from rich import print
from win32com.client import Dispatch

from Compare_Image import CompareImage
from helper import Helper
from var import *


class PowerPoint(Helper, CompareImage):

    def __init__(self):
        self.create_project_dirs()
        self.copy_for_test()
        self.compare(list_file_names_doc_from_compare,
                     extension_from,
                     extension_to)

    @staticmethod
    def opener_power_point(path_for_open, file_name):
        presentation = Dispatch("PowerPoint.application")
        presentation = presentation.Presentations.Open(f'{path_for_open}{file_name}', ReadOnly=True)
        slide_count = len(presentation.Slides)
        # print(presentation)
        print(slide_count)
        return slide_count, presentation

    def get_screenshot(self, file_name, path_to_save_screen):
        print(f'[bold green]In test[/bold green] {file_name}')
        slide_count, presentation = self.opener_power_point(path_to_folder_for_test, file_name)
        self.grab(path_to_save_screen, file_name, slide_count)
        presentation.Close()
        # sb.call(f'powershell.exe kill -Name POWERPNT', shell=True)
        return slide_count

    def compare(self, list_of_files, from_extension, to_extension):
        for file_name in list_of_files:
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            if to_extension == file_name.split('.')[-1]:
                slide_count_after = self.get_screenshot(file_name, path_to_tmpimg_after_conversion)
                slide_count_before = self.get_screenshot(file_name_from, path_to_tmpimg_befor_conversion)
                # sb.call(f'powershell.exe kill -Name POWERPNT', shell=True)

                if slide_count_after != slide_count_before:
                    self.copy(f'{custom_path_to_document_to}{file_name}', f'{path_to_errors_file}{file_name}')
                    self.copy(f'{custom_path_to_document_from}{file_name_from}', f'{path_to_errors_file}{file_name}')
                    print('[bold red]SLIDE COUNT DIFFERENT[/bold red]')
                else:
                    CompareImage()

        pass
