import os
import subprocess as sb
from multiprocessing import Process

from tqdm import tqdm

from config import *
from libs.functional.documents.doc_to_docx_image_compare import WordCompareImg
from libs.functional.documents.doc_to_docx_statistic_compare import Word
from libs.functional.presentation.ppt_to_pptx_compare import PowerPoint
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage
from libs.functional.spreadsheets.xls_to_xlsx_statistic_compare import Excel
from libs.helpers.get_error import run_get_errors_pp, run_get_error_exel
from libs.helpers.helper import Helper


def doc_docx_compare_statistic():
    for execution_time in tqdm(range(1)):
        error_processing = Process(target=run_get_errors_pp)
        error_processing.start()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        Word()
        error_processing.terminate()


def run_doc_docx_compare_image(list=False):
    for execution_time in tqdm(range(1)):
        helper = Helper('doc', 'docx')
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        if list:
            WordCompareImg(list_of_file_names, helper)
        else:
            WordCompareImg(os.listdir(helper.converted_doc_folder), helper)
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)


def run_doc_docx_full_test():
    for execution_time in tqdm(range(1)):
        helper = Helper('doc', 'docx')
        error_processing = Process(target=run_get_errors_pp)
        error_processing.start()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        Word(helper)
        WordCompareImg(os.listdir(helper.differences_statistic), helper)
        error_processing.terminate()


def run_ppt_pptx(list=False):
    for execution_time in tqdm(range(1)):
        helper = Helper('ppt', 'pptx')
        if list:
            PowerPoint(list_of_file_names, helper)
        else:
            PowerPoint(os.listdir(helper.converted_doc_folder), helper)


def run_xls_xlsx_compare_image(list=False):
    for i in tqdm(range(1)):
        helper = Helper('xls', 'xlsx')
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        if list:
            ExcelCompareImage(list_of_file_names, helper)
        else:
            ExcelCompareImage(os.listdir(helper.converted_doc_folder), helper)
        error_processing.terminate()


def run_xls_xlsx_compare_statistic():
    for execution_time in tqdm(range(1)):
        helper = Helper('xls', 'xlsx')
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        Excel(os.listdir(helper.converted_doc_folder), helper)
        error_processing.terminate()


def run_xls_xlsx_full():
    for execution_time in tqdm(range(1)):
        helper = Helper('xls', 'xlsx')
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        Excel(os.listdir(helper.converted_doc_folder), helper)
        ExcelCompareImage(helper.differences_statistic, helper)
        error_processing.terminate()
