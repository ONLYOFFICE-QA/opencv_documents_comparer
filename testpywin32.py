import json
import os

from win32com.client import Dispatch

# open Word
word = Dispatch('Word.Application')
word.Visible = False
word = word.Documents.Open(os.getcwd() + '/data/files/Alimentaire_Etude_Planning_StrategiqueKM_2010.docx')

# get number_of_pages of sheets
word.Repaginate()

statistics_word = {
    'num_of_sheets': f'{word.ComputeStatistics(2)}',
    'number_of_lines': f'{word.ComputeStatistics(1)}',
    'word_count': f'{word.ComputeStatistics(0)}',
    'number_of_characters_without_spaces': f'{word.ComputeStatistics(3)}',
    'number_of_characters_with_spaces': f'{word.ComputeStatistics(5)}',
    'number_of_abzad': f'{word.ComputeStatistics(4)}',
}
with open('configure.json', 'w') as f:
    json.dump(statistics_word, f)

print(statistics_word['num_of_sheets'])
word.Close()
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
