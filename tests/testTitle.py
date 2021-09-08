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


