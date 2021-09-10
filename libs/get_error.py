from time import sleep

import win32con
import win32gui
import pyautogui as pg
import subprocess as sb

ClassName = []

# coordinate = []
#
#
# def get_coord(hwnd, ctx):
#     if win32gui.IsWindowVisible(hwnd):
#         if win32gui.GetClassName(hwnd) == 'OpusApp':
#             win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
#             win32gui.SetForegroundWindow(hwnd)
#             coordinate.clear()
#             coordinate.append(win32gui.GetWindowRect(hwnd))
#     pass


def get_windows_title(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == '#32770' or win32gui.GetClassName(hwnd) == 'bosa_sdm_msword':
            # hwnd = win32gui.FindWindow(None, "Telegram (15125)")
            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
            win32gui.SetForegroundWindow(hwnd)

            ClassName.clear()
            ClassName.append(win32gui.GetClassName(hwnd))
            ClassName.append(win32gui.GetWindowText(hwnd))

        # elif win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) == 'Word':
        #     ClassName.clear()
        #     ClassName.append(win32gui.GetClassName(hwnd))
        #     ClassName.append(win32gui.GetWindowText(hwnd))


def run_get_pr():
    while True:
        win32gui.EnumWindows(get_windows_title, ClassName)
        sleep(0.2)
        if ClassName:
            check()
        # print(ClassName)
        # ClassName.clear()


def check():
    if ClassName[0] == '#32770':
        print(ClassName)
        if ClassName[1] == 'Microsoft Word':
            print(ClassName[1])
            pg.press('left')
            pg.press('enter')
            ClassName.clear()

        elif ClassName[1] == 'Microsoft Visual Basic for Applications':
            print(ClassName[1])
            pg.press('enter')
            ClassName.clear()

        elif ClassName[1] == 'Удаление нескольких элементов':
            print(ClassName[1])
            # pg.press('enter')
            ClassName.clear()

        elif ClassName[1] == 'Сохранить как':
            print(ClassName[1])
            sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
            ClassName.clear()

    elif ClassName[0] == 'bosa_sdm_msword':
        print(ClassName[1])
        if ClassName[1] == 'Преобразование файла':
            pg.press('enter')
            ClassName.clear()

        elif ClassName[1] == 'Пароль':
            pg.press('tab')
            pg.press('enter')
            ClassName.clear()

        elif ClassName[1] == 'Показать исправления':
            sleep(1)
            pg.press('tab', presses=3)
            pg.press('enter')
            ClassName.clear()
            pass

    # elif ClassName[0] == 'OpusApp':
    #     if ClassName[1] == 'Word':
    #         ClassName.clear()
    #         pass

# run_get_pr()
