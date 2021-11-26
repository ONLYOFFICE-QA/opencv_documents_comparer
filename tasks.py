# -*- coding: utf-8 -*-
import os

import pyperclip as pc
from invoke import task
from tqdm import tqdm

from config import *
from libs.functional.documents.doc_to_docx_image_compare import WordCompareImg
from libs.functional.presentation.ppt_to_pptx_compare import PowerPoint
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage
from libs.helpers.compare_image import CompareImage
from libs.helpers.helper import Helper


@task(name="doc-docx")
def run_doc_docx(c, full=False, st=False, ls=False, df=False, cl=False):
    for execution_time in tqdm(range(1)):
        word = WordCompareImg()
        if st:
            word.run_compare_word_statistic(os.listdir(word.helper.converted_doc_folder))
        elif full:
            word.run_compare_word_statistic(os.listdir(word.helper.converted_doc_folder))
            word.run_compare_word(os.listdir(word.helper.differences_statistic))
        elif ls:
            word.run_compare_word(list_of_file_names)
        elif cl:
            list_of_files = pc.paste()
            list_of_files = list_of_files.split("\n")
            word.run_compare_word(list_of_files)
        elif df:
            word.run_compare_word(os.listdir(word.helper.differences_statistic))
        else:
            word.run_compare_word(os.listdir(word.helper.converted_doc_folder))


@task(name="ppt-pptx")
def run_ppt_pptx(c, ls=False, cl=False):
    power_point = PowerPoint()
    if ls:
        power_point.run_compare_pp(list_of_file_names)
    elif cl:
        list_of_files = pc.paste()
        list_of_files = list_of_files.split("\n")
        power_point.run_compare_pp(list_of_files)
    else:
        power_point.run_compare_pp(os.listdir(power_point.helper.converted_doc_folder))


@task(name="xls_xlsx")
def run_xls_xlsx(c, full=False, st=False, ls=False, cl=False):
    for execution_time in tqdm(range(1)):
        excel = ExcelCompareImage()
        if full:
            excel.run_compare_exel_statistic(os.listdir(excel.helper.converted_doc_folder))
            excel.run_compare_excel_img(excel.helper.differences_statistic)
        elif ls:
            excel.run_compare_excel_img(list_of_file_names)
        elif cl:
            list_of_files = pc.paste()
            list_of_files = list_of_files.split("\n")
            excel.run_compare_excel_img(list_of_files)
        elif st:
            excel.run_compare_exel_statistic(os.listdir(excel.helper.converted_doc_folder))
        else:
            excel.run_compare_excel_img(os.listdir(excel.helper.converted_doc_folder))


# for test
@task(name="cmp")
def compare_image(c):
    helper = Helper('doc', 'docx')
    CompareImage('file.docx', helper)
