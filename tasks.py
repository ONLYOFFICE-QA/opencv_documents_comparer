# -*- coding: utf-8 -*-
import os

import pyperclip as pc
from invoke import task
from tqdm import tqdm

from config import *
from libs.functional.documents.doc_to_docx_image_compare import DocDocxCompareImg
from libs.functional.documents.doc_to_docx_statistic_compare import DocDocxStatisticsCompare
from libs.functional.documents.rtf_to_docx_image_compare import RtfDocxCompareImg
from libs.functional.presentation.odp_to_pptx_compare import OdpPptxCompare
from libs.functional.presentation.ppt_to_pptx_compare import PptPptxCompareImg
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage
from libs.openers.opener_docx_with_ms_word import OpenerDocx
from libs.openers.opener_pptx_with_ms_power_point import OpenerPptx


@task(name="doc-docx")
def run_doc_docx(c, full=False, st=False, ls=False, df=False, cl=False):
    for execution_time in tqdm(range(1)):
        if st:
            statistics_compare = DocDocxStatisticsCompare()
            statistics_compare.run_compare_word_statistic(os.listdir(statistics_compare.helper.converted_doc_folder))
        elif full:
            img_compare = DocDocxCompareImg()
            statistics_compare = DocDocxStatisticsCompare()
            statistics_compare.run_compare_word_statistic(os.listdir(img_compare.helper.converted_doc_folder))
            img_compare.run_compare_word(os.listdir(img_compare.helper.differences_statistic))
        elif ls:
            img_compare = DocDocxCompareImg()
            img_compare.run_compare_word(list_of_file_names)
        elif cl:
            img_compare = DocDocxCompareImg()
            list_of_files = pc.paste()
            list_of_files = list_of_files.split("\n")
            img_compare.run_compare_word(list_of_files)
        elif df:
            img_compare = DocDocxCompareImg()
            img_compare.run_compare_word(os.listdir(img_compare.helper.differences_statistic))
        else:
            img_compare = DocDocxCompareImg()
            img_compare.run_compare_word(os.listdir(img_compare.helper.converted_doc_folder))


@task(name="rtf-docx")
def run_rtf_docx(c, ls=False, cl=False):
    for execution_time in tqdm(range(1)):
        img_compare = RtfDocxCompareImg()
        if ls:
            img_compare.run_compare_rtf_docx(list_of_file_names)
        elif cl:
            list_of_files = pc.paste()
            list_of_files = list_of_files.split("\n")
            img_compare.run_compare_rtf_docx(list_of_files)
        else:
            img_compare.run_compare_rtf_docx(os.listdir(img_compare.word.helper.converted_doc_folder))


@task(name="ppt-pptx")
def run_ppt_pptx(c, ls=False, cl=False):
    power_point = PptPptxCompareImg()
    if ls:
        power_point.run_compare_pp(list_of_file_names)
    elif cl:
        list_of_files = pc.paste()
        list_of_files = list_of_files.split("\n")
        power_point.run_compare_pp(list_of_files)
    else:
        power_point.run_compare_pp(os.listdir(power_point.helper.converted_doc_folder))


@task(name="odp-pptx")
def run_odp_pptx(c, ls=False, cl=False):
    comparer = OdpPptxCompare()
    if ls:
        comparer.run_compare_odp_pptx(list_of_file_names)
    elif cl:
        list_of_files = pc.paste()
        list_of_files = list_of_files.split("\n")
        comparer.run_compare_odp_pptx(list_of_files)
    else:
        comparer.run_compare_odp_pptx(os.listdir(comparer.helper.converted_doc_folder))


@task(name="xls_xlsx")
def run_xls_xlsx(c, full=False, st=False, ls=False, cl=False):
    for execution_time in tqdm(range(1)):
        excel = ExcelCompareImage()
        if full:
            excel.run_compare_excel_statistic(os.listdir(excel.helper.converted_doc_folder))
            excel.run_compare_excel_img(excel.helper.differences_statistic)
        elif ls:
            excel.run_compare_excel_img(list_of_file_names)
        elif cl:
            list_of_files = pc.paste()
            list_of_files = list_of_files.split("\n")
            excel.run_compare_excel_img(list_of_files)
        elif st:
            excel.run_compare_excel_statistic(os.listdir(excel.helper.converted_doc_folder))
        else:
            excel.run_compare_excel_img(os.listdir(excel.helper.converted_doc_folder))


@task(name="opener_pptx")
def opener_pptx(c, odp=False, ppt=False):
    if odp:
        opener = OpenerPptx('odp')
        opener.run_opener(os.listdir(opener.helper.converted_doc_folder))
    elif ppt:
        opener = OpenerPptx('ppt')
        opener.run_opener(os.listdir(opener.helper.converted_doc_folder))
    else:
        opener = OpenerPptx('ppt')
        opener.run_opener(os.listdir(opener.helper.converted_doc_folder))
        opener = OpenerPptx('odp')
        opener.run_opener(os.listdir(opener.helper.converted_doc_folder))


@task(name="opener_docx")
def opener_docx(c, doc=False, rtf=False):
    if doc:
        opener = OpenerDocx('doc')
        opener.run_opener_word(os.listdir(opener.helper.converted_doc_folder))
    elif rtf:
        opener = OpenerDocx('rtf')
        opener.run_opener_word(os.listdir(opener.helper.converted_doc_folder))
    else:
        opener = OpenerDocx('doc')
        opener.run_opener_word(os.listdir(opener.helper.converted_doc_folder))
        opener = OpenerDocx('rtf')
        opener.run_opener_word(os.listdir(opener.helper.converted_doc_folder))


@task(name="compare")
def full_test(c):
    word = DocDocxCompareImg()
    word.run_compare_word(os.listdir(word.helper.converted_doc_folder))
    power_point = PptPptxCompareImg()
    power_point.run_compare_pp(os.listdir(power_point.helper.converted_doc_folder))
    excel = ExcelCompareImage()
    excel.run_compare_excel_img(os.listdir(excel.helper.converted_doc_folder))
