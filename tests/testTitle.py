# title = win32gui.GetWindowText(hwnd)
# print(title)
#
# def winEnumHandler(hwnd, ctx):
#     if win32gui.IsWindowVisible(hwnd):
#         if win32gui.GetWindowText(hwnd) == 'Word' and win32gui.GetWindowText(hwnd) == 'Пароль':
#             print(hex(hwnd), win32gui.GetWindowText(hwnd))
#         elif win32gui.GetWindowText(
#                 hwnd) == 'Alimentaire_Etude_Planning_StrategiqueKM_2010.docx [Режим ограниченной функциональности] - Word':
#             print(hex(hwnd), win32gui.GetClassName(hwnd))
import win32con
import win32gui

def get_windows_title(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        # if win32gui.GetClassName(hwnd) == '#32770':
        #     ClassName.append(win32gui.GetClassName(hwnd))
        #     ClassName.append(win32gui.GetWindowText(hwnd))
        #
        # # elif win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) != 'Word':
        # #     ClassName.append(win32gui.GetClassName(hwnd))
        # #     ClassName.append(win32gui.GetWindowText(hwnd))
        # else:
            print(f'class name  {win32gui.GetClassName(hwnd)}')
            print(f'Text name {win32gui.GetWindowText(hwnd)}')

ClassName = []
win32gui.EnumWindows(get_windows_title, ClassName)
print(ClassName)
print(ClassName)

# def window_get(window=None, class_name:str=None)->int:
#     ''' Returns hwnd. If window is not specified then
#         finds foreground window.
#     '''
#     if isinstance(window, str):
#         return win32gui.FindWindow(class_name, window)
#     elif isinstance(window, int):
#         return window
#     elif not window and class_name:
#         return win32gui.FindWindow(class_name, window)
#     else:
#         return win32gui.GetForegroundWindow()
#
#
# hwnd = window_get(class_name='bosa_sdm_msword')
# win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
# print(hwnd)
# #
hwnd = win32gui.FindWindow(None, "Telegram (15125)")
win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
win32gui.SetForegroundWindow(hwnd)

