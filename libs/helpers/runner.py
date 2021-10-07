import os
import subprocess as sb
from multiprocessing import Process

from tqdm import tqdm

from config import *
from libs.functional.documents.doc_to_docx_image_compare import WordCompareImg
from libs.functional.documents.doc_to_docx_statistic_compare import Word
from libs.functional.presentation.ppt_to_pptx_compare import PowerPoint
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage
from libs.helpers.get_error import run_get_errors_pp, run_get_error_exel


def doc_docx_compare_statistic():
    for execution_time in tqdm(range(1)):
        word = Word()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        error_processing = Process(target=run_get_errors_pp)
        error_processing.start()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        word.run_compare_word_statistic(os.listdir(word.helper.converted_doc_folder))
        error_processing.terminate()



def run_doc_docx_compare_image(list_of_files=False):
    for execution_time in tqdm(range(1)):
        word_compare = WordCompareImg()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        if list_of_files:
            word_compare.run_compare_word(list_of_file_names)
        else:
            word_compare.run_compare_word(os.listdir(word_compare.helper.converted_doc_folder))
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)


def run_doc_docx_full_test():
    for execution_time in tqdm(range(1)):
        word = WordCompareImg()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        error_processing = Process(target=run_get_errors_pp)
        error_processing.start()
        sb.call(f'powershell.exe kill -Name WINWORD', shell=True)
        word.run_compare_word_statistic(os.listdir(word.helper.converted_doc_folder))
        word.run_compare_word(os.listdir(word.helper.differences_statistic))
        error_processing.terminate()


def run_ppt_pptx_compare(list_of_files=False):
    for execution_time in tqdm(range(1)):
        power_point = PowerPoint()
        sb.call(f'powershell.exe kill -Name POWERPNT', shell=True)
        if list_of_files:
            power_point.run_compare_pp(list_of_file_names)
        else:
            power_point.run_compare_pp(os.listdir(power_point.helper.converted_doc_folder))
        sb.call(f'powershell.exe kill -Name POWERPNT', shell=True)


def run_xls_xlsx_compare_image(list_of_files=False):
    for i in tqdm(range(1)):
        excel = ExcelCompareImage()
        sb.call(f'powershell.exe kill -Name EXCEL', shell=True)
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        if list_of_files:
            excel.run_compare_excel_img(list_of_file_names)
        else:
            excel.run_compare_excel_img(os.listdir(excel.helper.converted_doc_folder))
        error_processing.terminate()
        sb.call(f'powershell.exe kill -Name EXCEL', shell=True)


def run_xls_xlsx_compare_statistic():
    for execution_time in tqdm(range(1)):
        excel = ExcelCompareImage()
        sb.call(f'powershell.exe kill -Name EXCEL', shell=True)
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        excel.run_compare_exel_statistic(os.listdir(excel.helper.converted_doc_folder))
        error_processing.terminate()
        sb.call(f'powershell.exe kill -Name EXCEL', shell=True)


def run_xls_xlsx_full():
    for execution_time in tqdm(range(1)):
        excel = ExcelCompareImage()
        sb.call(f'powershell.exe kill -Name EXCEL', shell=True)
        error_processing = Process(target=run_get_error_exel)
        error_processing.start()
        excel.run_compare_exel_statistic(os.listdir(excel.helper.converted_doc_folder))
        excel.run_compare_excel_img(excel.helper.differences_statistic)
        error_processing.terminate()
        sb.call(f'powershell.exe kill -Name EXCEL', shell=True)
