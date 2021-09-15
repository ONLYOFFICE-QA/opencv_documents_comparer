import os

version = '6.4.1.8'
extension_from = 'xls'
extension_to = 'xlsx'

wait_for_open = 10  # время ожидания открытия
wait_for_press = 0.5  # время ожидания открытия

list_file_names_doc_from_compare = [
        'Appendix 01 Function List.xlsx'
]

custom_doc_to = f'd:/ProjectsAndVM/data_db/results/{version}_xmllint_{extension_to}_{extension_from}/'
# custom_doc_to = f'D:/documents/results/{version}_xmllint_{extension_to}_{extension_from}/'
# custom_doc_to = f'd:/ProjectsAndVM/data_db/results/{version}_xmllint_{extension_from}_{extension_to}/'
custom_doc_from = f'd:/ProjectsAndVM/data_db/files/{extension_from}/'
# custom_doc_from = f'D:/documents/files/{extension_from}/'

ms_office = 'C:/Program Files (x86)/Microsoft Office/root/Office16/'
word = 'WINWORD.EXE'
power_point = 'POWERPNT.EXE'
exel = 'EXCEL.EXE'

# static data
path_to_project = os.getcwd()
path_to_data = path_to_project + '/data/'
path_to_folder_for_test = path_to_project + f'/data/{version}_{extension_from}_{extension_to}/'
path_to_errors_file = f'{path_to_folder_for_test}errors/'
path_to_errors_sim_file = f'{path_to_folder_for_test}errors_sim/'
path_to_not_tested_file = f'{path_to_folder_for_test}not_tested/'
path_to_result = f'{path_to_folder_for_test}result/'

# static tmp
path_to_compare_files = path_to_project + '/data/files/'
path_to_tmp = path_to_project + '/data/tmp/'
tmp_after = path_to_project + '/data/tmp/after/'
tmp_befor = path_to_project + '/data/tmp/before/'

path_to_temp_in_test = path_to_project + '/data/tmp/in_test/'

# папки исходных30 файлов должны называться как расширения,
# например если расширение из котрого мы конвертируем doc,
# то и папка где лежат эти файлы должна называться doc
