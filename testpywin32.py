import os

from win32com.client import Dispatch

# open Word
word = Dispatch('Word.Application')
word.Visible = False
word = word.Documents.Open(os.getcwd() + '/20100416005850549.doc')

# get number of sheets
word.Repaginate()
num_of_sheets = word.ComputeStatistics(2)
print(num_of_sheets)
word.Close()


def presentation_slide_count(filename):
    application = Dispatch("PowerPoint.application")
    presentation = application.Presentations.Open(filename)
    slide_count = len(presentation.Slides)
    presentation.Close()
    return slide_count
