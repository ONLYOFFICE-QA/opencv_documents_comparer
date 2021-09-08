from time import sleep

import win32gui
import pyautogui as pg
import subprocess as sb

ClassName = []


def get_windows_title(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == '#32770' and win32gui.GetWindowText(
                hwnd) == 'Microsoft Word':
            # pg.press('left')
            pg.press('enter')

        elif win32gui.GetClassName(hwnd) == '#32770':
            pg.press('enter')
            # sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
            # ClassName.append(win32gui.GetClassName(hwnd))
            # ClassName.append(win32gui.GetWindowText(hwnd))

        elif win32gui.GetClassName(hwnd) == 'bosa_sdm_msword':
            ClassName.append(win32gui.GetClassName(hwnd))
            ClassName.append(win32gui.GetWindowText(hwnd))
            sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
            pass

        elif win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) != 'Word':
            pass
            # ClassName.append(win32gui.GetClassName(hwnd))
            # ClassName.append(win32gui.GetWindowText(hwnd))


def run_get_pr():
    while True:
        win32gui.EnumWindows(get_windows_title, ClassName)
        sleep(0.2)
        # print(ClassName)
        ClassName.clear()
