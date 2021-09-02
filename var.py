import os

version = '6.4.1.8'
extension_from = 'doc'
extension_to = 'docx'

wait_for_open = 3  # время ожидания открытия

list_file_names_doc_from_compare = [
    '1267452368_Tornedalsradets_Arsmote_2009_protokoll.docx',
    'Alimentaire_Etude_Planning_StrategiqueKM_2010.docx'
]

path_to_project = os.getcwd()
custom_path_to_document_to = f'd:/ProjectsAndVM/data_db/results/{version}_xmllint_{extension_to}_{extension_from}/'
custom_path_to_document_from = f'd:/ProjectsAndVM/data_db/files/{extension_from}/'
path_to_compare_files = path_to_project + '/data/files/'
path_to_result = path_to_project + '/data/result/'
path_to_tmpimg_after_conversion = path_to_project + '/data/tmp/after/'
path_to_tmpimg_befor_conversion = path_to_project + '/data/tmp/before/'

# папки исходных файлов должны называться как расширения,
# например если расширение из котрого мы конвертируем doc,
# то и папка где лежат эти файлы должна называться doc
