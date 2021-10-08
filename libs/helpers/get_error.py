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
            check_errors_exel()


def check_errors_exel():
    if errors[0] == '#32770':
        if errors[1] == 'Microsoft Visual Basic':
            sb.call(["TASKKILL", "/IM", "EXCEL.EXE", "/t", "/f"], shell=True)
            errors.clear()
        elif errors[1] == 'Удаление нескольких элементов':
            errors.clear()

        elif errors[1] == 'Microsoft Excel':
            log.info('Microsoft Excel')
            pg.press('left')
            pg.press('enter')
            sleep(1)
            pg.press('enter')
            errors.clear()

        elif errors[1] == 'Monopoly':
            log.info('Monopoly')
            pg.press('enter')
            errors.clear()

    elif errors[0] == 'ThunderDFrame':
        if errors[1] == 'Functions List':
            print('Functions List')
            pg.hotkey('alt', 'f4')
            errors.clear()

        elif errors[1] == 'Select Players and Times':
            log.info('Select Players and Times')
            pg.press('tab', presses=6, interval=0.2)
            pg.press('enter', interval=0.2)
            errors.clear()

    elif errors[0] == 'NUIDialog':
        if errors[1] == 'Microsoft Excel' or errors[1] == 'Microsoft Excel - проверка совместимости':
            print('Microsoft Excel')
            pg.press('enter')
            errors.clear()


def run_get_errors_pp():
    while True:
        win32gui.EnumWindows(get_windows_title, errors)
        sleep(0.2)
        if errors:
            check_pp()


def check_pp():
    if errors[0] == '#32770':
        print(errors)
        if errors[1] == 'Microsoft Word':
            print(errors[1])
            pg.press('left')
            pg.press('enter')
            errors.clear()

        elif errors[1] == 'Microsoft Visual Basic for Applications':
            print(errors[1])
            pg.press('enter')
            errors.clear()

        elif errors[1] == 'Удаление нескольких элементов':
            print(errors[1])
            # pg.press('enter')
            errors.clear()

        elif errors[1] == 'Сохранить как':
            print(errors[1])
            sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
            errors.clear()

    elif errors[0] == 'bosa_sdm_msword':
        print(errors[1])
        if errors[1] == 'Преобразование файла':
            pg.press('enter')
            errors.clear()

        elif errors[1] == 'Пароль':
            pg.press('tab')
            pg.press('enter')
            errors.clear()

        elif errors[1] == 'Показать исправления':
            sleep(1)
            pg.press('tab', presses=3)
            pg.press('enter')
            errors.clear()
            pass

