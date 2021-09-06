import json
import subprocess as sb

import win32gui
from rich import print
from rich.progress import track
from win32com.client import Dispatch

from helper import Helper as operation
from var import *


def get_word_metadata(word):
    statistics_word = {
        'num_of_sheets': f'{word.ComputeStatistics(2)}',
        'number_of_lines': f'{word.ComputeStatistics(1)}',
        'word_count': f'{word.ComputeStatistics(0)}',
        'number_of_characters_without_spaces': f'{word.ComputeStatistics(3)}',
        'number_of_characters_with_spaces': f'{word.ComputeStatistics(5)}',
        'number_of_abzad': f'{word.ComputeStatistics(4)}',
    }
    return statistics_word


def get_windows_title(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) != 'Word':
            print(f'class name  {win32gui.GetClassName(hwnd)}')
            print(f'Text name {win32gui.GetWindowText(hwnd)}')
            print(f'GetCapture {win32gui.GetCapture()}')
            print(f'DesktopWindow {win32gui.GetDesktopWindow()}')
            print(f'ForegroundWindow {win32gui.GetForegroundWindow()}')
            print(f'GetMenu {win32gui.GetMenu(hwnd)}')
            print(f'GetMenuItemCount {win32gui.GetMenuItemCount(hwnd)}')
            print(f'GetParent {win32gui.GetParent(hwnd)}')
            # print(f'G {win32gui.GetTextAlign(hwnd)}')
            print(f'GetTextColor {win32gui.GetTextColor(hwnd)}')
            # print(f'G {win32gui.GetTextFace(hwnd)}')
            # print(f'G {win32gui.GetTextMetrics(hwnd)}')
            print(f'GetWindowDC {win32gui.GetWindowDC(hwnd)}')
            print(f'GetWindowPlacement {win32gui.GetWindowPlacement(hwnd)}')  # получает позицию окна
            print(f'GetWindowRect {win32gui.GetWindowRect(hwnd)}')  # получает позицию окна
            print(f'GetWindowTextLength {win32gui.GetWindowTextLength(hwnd)}')
            # print(f'G {win32gui.GetDlgItemText(hwnd)}')
        else:
            print(f'class name  {win32gui.GetClassName(hwnd)}')
            print(f'Text name {win32gui.GetWindowText(hwnd)}')

    win32gui.EnumWindows(get_windows_title, None)


def word_opener(file_name, path_for_open):
    # file_name_for_test = operation.preparing_files(file_name)
    # operation.copy(f'{path_for_open}{file_name}',
    #                f'{path_to_temp_in_test}{file_name_for_test}')
    word = Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    # word = word.Documents.Open( f'{path_to_temp_in_test}{file_name_for_test}')
    try:
        word = word.Documents.Open(f'{path_for_open}{file_name}', None, True)
        # word.Repaginate()
        statistics_word = get_word_metadata(word)
    except Exception:
        print('[bold red]NOT TESTED!!![/bold red]')
        operation.copy(f'{path_for_open}{file_name}', f'{path_to_not_tested_file}{file_name}')
        statistics_word = {}

    # word.Close()
    sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
    return statistics_word


def open_document_and_compare(list_of_files, from_extension, to_extension):
    for file_name in track(list_of_files, description='[bold blue]Comparing... [/bold blue]'):
        if to_extension == file_name.split('.')[-1]:
            print(f'[bold green]In test[/bold green] {file_name}')
            statistics_word_after = word_opener(file_name, custom_path_to_document_to)
            file_name_from = file_name.replace(f'.{to_extension}', f'.{from_extension}')
            print(f'[bold green]In test[/bold green] {file_name_from}')
            statistics_word_before = word_opener(file_name_from, custom_path_to_document_from)

            modified = dict_compare(statistics_word_after, statistics_word_before)

            if modified != {}:
                print(modified)
                operation.copy(f'{custom_path_to_document_to}{file_name}',
                               f'{path_to_errors_file}{file_name}')

                operation.copy(f'{custom_path_to_document_from}{file_name_from}',
                               f'{path_to_errors_file}{file_name_from}')
                with open(f'{path_to_errors_file}{file_name}_difference.json', 'w') as f:
                    json.dump(modified, f)
                pass
    operation.delete(path_to_temp_in_test)
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(["taskkill", "/IM", "WINWORD.EXE"])
    # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)


def create_project_dirs():
    operation.create_dir(path_to_tmpimg_befor_conversion)
    operation.create_dir(path_to_tmpimg_after_conversion)
    operation.create_dir(path_to_result)
    operation.create_dir(path_to_temp_in_test)


def dict_compare(d1, d2):
    d1_keys = set(d1.keys())
    d2_keys = set(d2.keys())
    shared_keys = d1_keys.intersection(d2_keys)
    modified = {o: (f'Before {d1[o]}', f'After {d2[o]}') for o in shared_keys if d1[o] != d2[o]}
    return modified
