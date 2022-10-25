# -*- coding: utf-8 -*-
from framework.telegram import Telegram
from libs.functional.presentation.odp_to_pptx_compare import OdpPptxCompare
from libs.functional.presentation.ppt_to_pptx_compare import PptPptxCompareImg
from libs.functional.spreadsheets.xls_to_xlsx_image_compare import ExcelCompareImage
from libs.functional.spreadsheets.xls_to_xlsx_statistic_compare import StatisticCompare
from libs.functional.documents.doc_to_docx_image_compare import DocDocxCompareImg
from libs.functional.documents.doc_to_docx_statistic_compare import DocDocxStatisticsCompare
from libs.functional.documents.rtf_to_docx_image_compare import RtfDocxCompareImg
from libs.helpers.document_helper import DocHelper
from libs.openers.opener_docx_with_ms_word import OpenerDocx
from libs.openers.opener_odp_with_libre_office import OpenerOdp
from libs.openers.opener_ods_with_libre_office import OpenerOds
from libs.openers.opener_odt_with_libre_office import OpenerOdt
from libs.openers.opener_pptx_with_ms_power_point import OpenerPptx
from libs.openers.opener_xlsx_with_ms_excell import OpenerXlsx
from management import task


@task()
def doc_docx(c, st=False, ls=False, df=False, cl=False):
    doc_helper = DocHelper(source_extension='doc', converted_extension='docx')
    files_array = doc_helper.get_file_array(ls=ls, df=df, cl=cl)
    comparer = DocDocxCompareImg(doc_helper) if not st else DocDocxStatisticsCompare(doc_helper)
    comparer.run_compare(files_array) if not st else comparer.run_compare_statistic(files_array)
    Telegram.send_message('doc-docx comparison completed')


@task()
def rtf_docx(c, ls=False, cl=False):
    doc_helper = DocHelper(source_extension='rtf', converted_extension='docx')
    comparer = RtfDocxCompareImg(doc_helper)
    comparer.run_compare(doc_helper.get_file_array(ls=ls, cl=cl))
    Telegram.send_message('rtf-docx comparison completed')


@task()
def ppt_pptx(c, ls=False, cl=False):
    doc_helper = DocHelper(source_extension='ppt', converted_extension='pptx')
    comparer = PptPptxCompareImg(doc_helper)
    comparer.run_compare(doc_helper.get_file_array(ls=ls, cl=cl))
    # Telegram.send_message('ppt-pptx comparison completed')


@task()
def odp_pptx(c, ls=False, cl=False):
    doc_helper = DocHelper(source_extension='odp', converted_extension='pptx')
    comparer = OdpPptxCompare(doc_helper)
    comparer.run_compare(doc_helper.get_file_array(ls=ls, cl=cl))
    Telegram.send_message('odp-pptx comparison completed')


@task()
def xls_xlsx(c, st=False, ls=False, cl=False):
    doc_helper = DocHelper(source_extension='xls', converted_extension='xlsx')
    files_array = doc_helper.get_file_array(ls=ls, cl=cl)
    comparer = ExcelCompareImage(doc_helper) if not st else StatisticCompare(doc_helper)
    comparer.run_compare(files_array) if not st else comparer.run_compare_statistic(files_array)
    Telegram.send_message('xls-xlsx comparison completed')


def run_opener_pptx(doc_helper, ls):
    opener = OpenerPptx(doc_helper)
    opener.run_opener(doc_helper.get_file_array(ls=ls))
    doc_helper.create_massage_for_tg(opener.errors_files_when_opening, ls=ls)


@task
def opener_pptx(c, odp=False, ppt=False, ls=False):
    if odp:
        run_opener_pptx(DocHelper(source_extension='odp', converted_extension='pptx'), ls)
    elif ppt:
        run_opener_pptx(DocHelper(source_extension='ppt', converted_extension='pptx'), ls)
    else:
        opener_pptx(c, ppt=True, ls=ls)
        opener_pptx(c, odp=True, ls=ls)
        Telegram.send_message('Doc=>Docx and Rtf => Docx opening check completed')


def run_opener_docx(doc_helper, ls):
    opener = OpenerDocx(doc_helper)
    opener.run_opener(doc_helper.get_file_array(ls=ls))
    doc_helper.create_massage_for_tg(opener.errors_files_when_opening, ls=ls)


@task
def opener_docx(c, doc=False, rtf=False, pdf=False, ls=False):
    if doc:
        run_opener_docx(DocHelper(source_extension='doc', converted_extension='docx'), ls)
    elif rtf:
        run_opener_docx(DocHelper(source_extension='rtf', converted_extension='docx'), ls)
    elif pdf:
        run_opener_docx(DocHelper(source_extension='pdf', converted_extension='docx'), ls)
    else:
        opener_docx(c, doc=True, ls=ls)
        opener_docx(c, rtf=True, ls=ls)
        Telegram.send_message('Doc=>Docx and Rtf=>Docx opening check completed')


def run_opener_xlsx(doc_helper, ls):
    opener = OpenerXlsx(doc_helper)
    opener.run_opener(doc_helper.get_file_array(ls=ls))
    doc_helper.create_massage_for_tg(opener.errors_files_when_opening, ls=ls)


@task
def opener_xlsx(c, xls=False, ods=False, ls=False):
    if xls:
        run_opener_xlsx(DocHelper(source_extension='xls', converted_extension='xlsx'), ls)
    elif ods:
        run_opener_xlsx(DocHelper(source_extension='ods', converted_extension='xlsx'), ls)
    else:
        opener_xlsx(c, xls=True, ls=ls)
        opener_xlsx(c, ods=True, ls=ls)
        Telegram.send_message('Xls=>Xlsx and Ods=>Xlsx opening check completed')


@task
def opener_odp(c, ls=False):
    doc_helper = DocHelper(source_extension='pptx', converted_extension='odp')
    opener = OpenerOdp(doc_helper)
    opener.run_opener(doc_helper.get_file_array(ls=ls))
    doc_helper.create_massage_for_tg(opener.errors_files_when_opening, ls=ls)


@task
def opener_odt(c, ls=False):
    doc_helper = DocHelper(source_extension='docx', converted_extension='odt')
    opener = OpenerOdt(doc_helper)
    opener.run_opener(doc_helper.get_file_array(ls=ls))
    doc_helper.create_massage_for_tg(opener.errors_files_when_opening, ls=ls)


@task
def opener_ods(c, ls=False):
    doc_helper = DocHelper(source_extension='xlsx', converted_extension='ods')
    opener = OpenerOds(doc_helper)
    opener.run_opener(doc_helper.get_file_array(ls=ls))
    doc_helper.create_massage_for_tg(opener.errors_files_when_opening, ls=ls)


@task
def opener_full(c):
    opener_pptx(c)
    opener_docx(c)
    opener_xlsx(c)
    opener_ods(c)
    opener_odp(c)
    opener_odt(c)
    Telegram.send_message('Full test of the openers completed')
