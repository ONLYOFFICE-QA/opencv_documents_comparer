from invoke import task

from runner import *


@task(name="doc-docx")
def doc_docx_statistic(im=False, st=False, l=False):
    if st:
        doc_docx_compare_statistic()
    elif im:
        if l:
            run_doc_docx_compare_image(list_of_file_names)
        else:
            run_doc_docx_compare_image(os.listdir(converted_doc_folder))
    else:
        run_doc_docx_full_test()


@task(name="ppt-pptx")
def run_ppt_pptx(full=False, l=False):
    if l:
        run_ppt_pptx_list()
    elif full:
        run_ppt_pptx_full()


@task(name="xls-xlsx")
def run_xls_xlsx(im=False, st=False, l=False):
    if im:
        if l:
            run_xls_xlsx_compare_image(list_of_file_names)
        else:
            run_xls_xlsx_compare_image(os.listdir(converted_doc_folder))
    elif st:
        run_xls_xlsx_compare_statistic()
    else:
        run_xls_xlsx_full()
