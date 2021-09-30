import os

version = '6.4.1.44'
extension_source = 'doc'
extension_converted = 'docx'

# Opening waiting time
wait_for_opening = 8
wait_for_press = 0.5

# To test from the names array
list_of_file_names = [
    'arturomurante_cv.docx',
    'Article_contes.docx'
]

# Path to source documents
source_doc_folder = f'd:/ProjectsAndVM/data_db/files/{extension_source}/'
# Path to converted documents
converted_doc_folder = f'd:/ProjectsAndVM/data_db/results/{version}_xmllint_{extension_source}_{extension_converted}/'

# office path
ms_office = 'C:/Program Files (x86)/Microsoft Office/root/Office16/'
word = 'WINWORD.EXE'
power_point = 'POWERPNT.EXE'
exel = 'EXCEL.EXE'

# static data
project_folder = os.getcwd()
data = project_folder + '/data/'
result_folder = f'{project_folder}/data/{version}_{extension_source}_{extension_converted}/'
differences_statistic = f'{result_folder}differences_statistic/'
differences_compare_image = f'{result_folder}differences_compare_image/'
untested_folder = f'{result_folder}untested/'
failed_source = f'{result_folder}failed_source/'
passed = f'{result_folder}passed/'

# static tmp
tmp_dir = project_folder + '/data/tmp/'
tmp_converted_image = project_folder + '/data/tmp/converted_image/'
tmp_source_image = project_folder + '/data/tmp/source_image/'
tmp_in_test = project_folder + '/data/tmp/in_test/'
