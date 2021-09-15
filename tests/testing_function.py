# def compare_img_wirh_imagemagic(before_conversion, after_conversion, image_name):
#     try:
#         folder = image_name.split('.')
#         folder = folder[0] + '_imgmagic'
#         os.mkdir(os.getcwd() + f'\\data/{folder}')
#     except Exception:
#         print(Exception, 'folder created')
#     command = f'magick compare -compose src {before_conversion} {after_conversion} data/{folder}/{image_name}.png'
#     sb.call(shlex.split(command), shell=True)
#     pass
#
#
# def grab(filename, path):
#     im = ImageGrab.grab()
#     im.save(path + filename + '.png', 'PNG')
#     return im

from win32com.client import Dispatch

from var import *


# # open Word
# word = Dispatch('Word.Application')
# word.Visible = False
# word = word.Documents.Open(os.getcwd() + '/20100416005850549.doc')
#
# # get number_of_pages of sheets
# word.Repaginate()
# num_of_sheets = word.ComputeStatistics(2)
# print(num_of_sheets)
# word.Close()


def presentation_slide_count(filename):
    presentation = Dispatch("PowerPoint.application")
    presentation.Visible = True
    presentation = presentation.Presentations.Open(filename)
    slide_count = len(presentation.Slides)
    # slide_count2 = len(presentation.Visible)
    print(presentation)
    print(slide_count)
    # print(slide_count2)
    presentation.Close()


presentation_slide_count(f'{custom_doc_to}2.pptx')
