# -*- coding: utf-8 -*-
import os

from tqdm import tqdm

from config import *
from libs.functional.documents.doc_to_docx_image_compare import WordCompareImg
from libs.functional.documents.doc_to_docx_statistic_compare import Word
from libs.functional.presentation.ppt_to_pptx_compare import PowerPoint
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage


def doc_docx_compare_statistic():
    for execution_time in tqdm(range(1)):
        word = Word()
        word.run_compare_word_statistic(os.listdir(word.helper.converted_doc_folder))


def run_doc_docx_compare_image(list_of_files=False, differences_statistic=False):
    for execution_time in tqdm(range(1)):
        word_compare = WordCompareImg()
        if list_of_files:
            word_compare.run_compare_word(list_of_file_names)

        elif differences_statistic:
            word_compare.run_compare_word(os.listdir(word_compare.helper.differences_statistic))
        else:
            word_compare.run_compare_word(os.listdir(word_compare.helper.converted_doc_folder))


def run_doc_docx_full_test():
    for execution_time in tqdm(range(1)):
        word = WordCompareImg()
        word.run_compare_word_statistic(os.listdir(word.helper.converted_doc_folder))
        word.run_compare_word(os.listdir(word.helper.differences_statistic))


def run_ppt_pptx_compare(list_of_files=False):
    for execution_time in tqdm(range(1)):
        power_point = PowerPoint()
        if list_of_files:
            power_point.run_compare_pp(list_of_file_names)
        else:
            power_point.run_compare_pp(os.listdir(power_point.helper.converted_doc_folder))


def run_xls_xlsx_compare_image(list_of_files=False):
    for i in tqdm(range(1)):
        excel = ExcelCompareImage()
        if list_of_files:
            excel.run_compare_excel_img(list_of_file_names)
        else:
            excel.run_compare_excel_img(os.listdir(excel.helper.converted_doc_folder))


def run_xls_xlsx_compare_statistic():
    for execution_time in tqdm(range(1)):
        excel = ExcelCompareImage()
        excel.run_compare_exel_statistic(os.listdir(excel.helper.converted_doc_folder))


def run_xls_xlsx_full():
    for execution_time in tqdm(range(1)):
        excel = ExcelCompareImage()
        excel.run_compare_exel_statistic(os.listdir(excel.helper.converted_doc_folder))
        excel.run_compare_excel_img(excel.helper.differences_statistic)
