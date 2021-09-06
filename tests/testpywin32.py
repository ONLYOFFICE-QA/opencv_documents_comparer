import os

from win32com.client import Dispatch

# open Word

# word.ReadOnly = True
# # word.DisplayAlerts = False
# print('1')
# # word = word.Documents.Open(os.getcwd() + '/data/02c.docx')
#
# # sb.Popen([r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE", f"{os.getcwd()}/data/02c.docx"])
#
# print(word)

# def errors():
#     def winEnumHandler(hwnd, ctx):
#         if win32gui.IsWindowVisible(hwnd):
#             if win32gui.GetClassName(hwnd) == 'bosa_sdm_msword'  and win32gui.GetWindowText(hwnd) != 'Пароль':
#                 pg.press('tab')
#                 pg.press('enter')
#                 return '1'
#
#
#
#             elif win32gui.GetClassName(hwnd) == 'OpusApp' and win32gui.GetWindowText(hwnd) != 'Word':
#                 print('opened')
#                 return '2'
#
#     win32gui.EnumWindows(winEnumHandler, None)

# get number_of_pages of sheets
word = Dispatch('Word.Application')
word.Visible = False

try:

    word = word.Documents.Open(os.getcwd() + '/data/240-0414.doc')
    statistics_word = {
        'num_of_sheets': f'{word.ComputeStatistics(2)}',
        'number_of_lines': f'{word.ComputeStatistics(1)}',
        'word_count': f'{word.ComputeStatistics(0)}',
        'number_of_characters_without_spaces': f'{word.ComputeStatistics(3)}',
        'number_of_characters_with_spaces': f'{word.ComputeStatistics(5)}',
        'number_of_abzad': f'{word.ComputeStatistics(4)}',
    }

except Exception:
    statistics_word = {}
    print(Exception)

# er = errors()
# print(er)
# word.Repaginate()

#
# with open('configure.json', 'w') as f:
#     json.dump(statistics_word, f)

print(statistics_word)

# word.Close()
# 1 - количество строк
# 0 - клличество слов
# 3 - знаков без пробелов
# 4 - абзадцев
# 5 - знаков с пробелами
#
# application = Dispatch("PowerPoint.application")
# presentation = application.Presentations.Open(os.getcwd() + '/data/errors/synopexgreentech.ppt')
# slide_count = len(presentation.Slides)
# print(slide_count)
# presentation.Close()
