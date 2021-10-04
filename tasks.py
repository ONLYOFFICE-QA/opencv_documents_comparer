from invoke import task

from libs.helpers.runner import *


@task(name="doc-docx")
def doc_docx_statistic(c, im=False, st=False, ls=False):
    if st:
        doc_docx_compare_statistic()
    elif im:
        if ls:
            run_doc_docx_compare_image(list=True)
        else:
            run_doc_docx_compare_image()
    else:
        run_doc_docx_full_test()


@task(name="ppt-pptx")
def run_ppt_pptx(c, full=False, ls=False):
    if ls:
        run_ppt_pptx(list=True)
    elif full:
        run_ppt_pptx()


@task(name="xls-xlsx")
def run_xls_xlsx(c, im=False, st=False, ls=False):
    if im:
        if ls:
            run_xls_xlsx_compare_image(list=True)
        else:
            run_xls_xlsx_compare_image()
    elif st:
        run_xls_xlsx_compare_statistic()
    else:
        run_xls_xlsx_full()
