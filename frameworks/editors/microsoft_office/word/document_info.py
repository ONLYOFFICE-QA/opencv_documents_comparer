# -*- coding: utf-8 -*-
from os.path import basename
from frameworks.StaticData import StaticData
from frameworks.decorators import async_processing
from host_tools import File, Window, HostInfo, Shell

from .handlers import WordEvents

if HostInfo().os == 'windows':
    from win32com.client import Dispatch


def handler():
    window_control = Window()
    while True:
        window_control.handle_events(WordEvents.handler)


@async_processing(target=handler)
class DocumentInfo:
    def __init__(self, file_path):
        self.tmp_dir = StaticData.tmp_dir_in_test
        self.word = StaticData.word
        self.tmp_file = File.make_tmp(file_path, self.tmp_dir)
        self.document_name = basename(file_path)
        self.word_app = Dispatch('Word.Application')
        self.word_app.Visible = False
        self.document = self.__open(self.tmp_file)

    def __del__(self):
        self.document.Close(False) if self.document else ...
        self.word_app.Quit()
        Shell.call(f"taskkill /im {self.word}")
        File.delete(self.tmp_file, stdout=False)

    def get(self):
        try:
            return {
                'page_amount': f'{self.document.ComputeStatistics(2)}',
                'line_amount': f'{self.document.ComputeStatistics(1)}',
                'word_count': f'{self.document.ComputeStatistics(0)}',
                'characters_without_spaces': f'{self.document.ComputeStatistics(3)}',
                'characters_with_spaces': f'{self.document.ComputeStatistics(5)}',
                'paragraph_amount': f'{self.document.ComputeStatistics(4)}',
            }
        except Exception as e:
            print(f"Exception when getting full information about document: {e}")

    def page_amount(self):
        try:
            return self.document.ComputeStatistics(2)
        except Exception as e:
            print(f"Exception when getting page amount: {e}")

    def __open(self, file_path: str):
        try:
            return self.word_app.Documents.Open(f'{file_path}', None, True)
        except Exception as e:
            return print(f"Exception {e} when open document: {self.document_name}")
