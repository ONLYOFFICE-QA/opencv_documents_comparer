# -*- coding: utf-8 -*-
import subprocess as sb
from time import sleep

import pyautogui as pg
import win32con
import win32gui

from libs.helpers.logger import *

errors = []


def get_windows_title(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == '#32770' \
                or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword' \
                or win32gui.GetClassName(hwnd) == 'ThunderDFrame' \
                or win32gui.GetClassName(hwnd) == 'NUIDialog':
            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
            win32gui.SetForegroundWindow(hwnd)

            errors.clear()
            errors.append(win32gui.GetClassName(hwnd))
            errors.append(win32gui.GetWindowText(hwnd))


def run_get_error_exel():
    while True:
        win32gui.EnumWindows(get_windows_title, errors)
        sleep(0.2)
        if errors:
            check_errors_excel(errors)
            errors.clear()


def check_errors_excel(array_of_errors):
    if array_of_errors[0] == '#32770':
        if array_of_errors[1] == 'Microsoft Visual Basic':
            sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)
        elif array_of_errors[1] == 'Удаление нескольких элементов':
            pass

        elif array_of_errors[1] == 'Microsoft Excel':
            log.info('Microsoft Excel')
            pg.press('left')
            pg.press('enter')
            sleep(1)
            pg.press('enter')

        elif array_of_errors[1] == 'Monopoly':
            log.info('Monopoly')
            pg.press('enter')

    elif array_of_errors[0] == 'ThunderDFrame':
        if array_of_errors[1] == 'Functions List':
            print('Functions List')
            pg.hotkey('alt', 'f4')

        elif array_of_errors[1] == 'Select Players and Times':
            log.info('Select Players and Times')
            pg.press('tab', presses=6, interval=0.2)
            pg.press('enter', interval=0.2)

    elif array_of_errors[0] == 'NUIDialog':
        if array_of_errors[1] == 'Microsoft Excel' or array_of_errors[1] == 'Microsoft Excel - проверка совместимости':
            print('Microsoft Excel')
            pg.press('enter')


def run_get_errors_word():
    while True:
        win32gui.EnumWindows(get_windows_title, errors)
        sleep(0.2)
        if errors:
            check_word(errors)
            errors.clear()


def check_word(array_of_errors):
    if array_of_errors[0] == '#32770':
        print(array_of_errors)
        if array_of_errors[1] == 'Microsoft Word':
            print(array_of_errors[1])
            pg.press('left')
            pg.press('enter')

        elif array_of_errors[1] == 'Microsoft Visual Basic for Applications':
            print(array_of_errors[1])
            pg.press('enter')

        elif array_of_errors[1] == 'Удаление нескольких элементов':
            print(array_of_errors[1])
            # pg.press('enter')

        elif array_of_errors[1] == 'Сохранить как':
            print(array_of_errors[1])
            sb.call(f'powershell.exe kill -Name WINWORD', shell=True)

    elif array_of_errors[0] == 'bosa_sdm_msword':
        print(array_of_errors[1])
        if array_of_errors[1] == 'Преобразование файла':
            pg.press('enter')

        elif array_of_errors[1] == 'Пароль':
            pg.press('tab')
            pg.press('enter')

        elif array_of_errors[1] == 'Показать исправления':
            sleep(1)
            pg.press('tab', presses=3)
            pg.press('enter')
