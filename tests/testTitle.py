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
import psutil
# from pywinauto import Desktop
#
# windows = Desktop(backend="uia").windows()
# print(windows)


def get_windows_title(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        if win32gui.GetClassName(hwnd) == '#32770':
            ClassName.append(win32gui.GetClassName(hwnd))
            ClassName.append(win32gui.GetWindowText(hwnd))

        elif win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) != 'Word':
            ClassName.append(win32gui.GetClassName(hwnd))
            ClassName.append(win32gui.GetWindowText(hwnd))


ClassName = []
win32gui.EnumWindows(get_windows_title, ClassName)
print(ClassName[0])
print(ClassName[1])

# #
# import win32process
# # Список процессов с именем файла notepad.exe:
# notepads = [item for item in psutil.process_iter() if item.name() == 'WINWORD.EXE']
# print(notepads)  # [<psutil.Process(pid=4416, name='notepad.exe') at 64362512>]
#
# # Просто pid первого попавшегося процесса с именем файла notepad.exe:
# pid = next(item for item in psutil.process_iter() if item.name() == 'WINWORD.EXE').pid
# # (вызовет исключение StopIteration, если Блокнот не запущен)
#
# print(pid)
# def enum_window_callback(hwnd, pid):
#     tid, current_pid = win32process.GetWindowThreadProcessId(hwnd)
#     if pid == current_pid and win32gui.IsWindowVisible(hwnd):
#         windows.append(hwnd)
#
# # pid = 4416  # pid уже получен на предыдущем этапе
# windows = []
#
# win32gui.EnumWindows(enum_window_callback, pid)
#
# # Выводим заголовки всех полученных окон
# print([win32gui.GetWindowText(item) for item in windows])