import os
import shutil
import subprocess as sb

from rich.progress import track


class Helper:
    # path insert with file name
    def copy(path_from, path_to):
        if os.path.exists(path_from):
            shutil.copyfile(path_from, path_to)
        else:
            print(f'Folder exists: {path_from}')

    def create_dir(path_to_dir):
        if not os.path.exists(path_to_dir):
            os.mkdir(path_to_dir)
        else:
            print(f'Folder exists: {path_to_dir}')

    # принимает к каталогу с фалвми которые нужно переименовать
    def rename_files(dir_for_rename_files):
        for file in track(os.listdir(dir_for_rename_files)):
            # print(os.listdir(dir_for_rename_files))
            extension = file.split('.')[-1]
            file_name = file.replace(f'.{extension}', '')
            file_name = file_name.replace(')', '_')
            file_name = file_name.replace('(', '_')
            file_name = file_name.replace(' ', '_')
            if file != file_name + f'.{extension}':
                # shutil.move(f'{files_for_test_path}{file}', f'{files_for_test_path}{folder_name}{file[-5:]}')
                os.rename(f'{dir_for_rename_files}{file}', f'{dir_for_rename_files}{file_name}.{extension}')

    def preparing_files(file_name):
        file_name_for_test = file_name.replace(')', '_')
        file_name_for_test = file_name_for_test.replace('(', '_')
        file_name_for_test = file_name_for_test.replace(' ', '_')
        return file_name_for_test

    def delete(what_delete):
        sb.call(f'powershell.exe rm {what_delete} -Force -Recurse', shell=True)
