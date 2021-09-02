import os
import shutil

path_list_from = 'd:\\ProjectsAndVM\data_db\\results\\6.5.0.1_xmllint_pptx_ppt\\errors\\'
path_files_from = 'd:\\ProjectsAndVM\\data_db\\results\\6.4.1.8_xmllint_pptx_ppt\\'
path_to_sours_files = 'D:\\ProjectsAndVM\\data_db\\files\\ppt\\'

path_to_copy = 'D:\\ProjectsAndVM\\conversion\CompareIMG\\data\\test_folder\\'

list_files = os.listdir(path_list_from)

for file in list_files:
    file_Extension = file.split('.')[-1]
    if file_Extension == 'pptx':
        # shutil.copyfile(path_files_from + file, path_to_copy + file)
        file = file.replace(f'{file_Extension}', 'ppt')  # Convert extension
        shutil.copyfile(path_to_sours_files + file, path_to_copy + file)
