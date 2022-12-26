# -*- coding: utf-8 -*-

import os
import json
import random
import shutil
import subprocess as sb
import codecs
import zipfile

import py7zr
from rich import print


class FileUtils:
    @staticmethod
    def read_json(path_to_json):
        with codecs.open(path_to_json, "r", "utf_8_sig") as file:
            return json.load(file)

    # path insert with file name
    @staticmethod
    def copy(path_from, path_to, silence=False):
        if not os.path.exists(path_from):
            return print(f"[red]copy warning: path not exist: {path_from}")
        if os.path.isdir(path_from):
            shutil.copytree(path_from, path_to)
        elif os.path.isfile(path_from):
            shutil.copyfile(path_from, path_to)
        if os.path.exists(path_to):
            return print(f'[green]Copied To:[/] {path_to}') if not silence else None
        return print(f'[red]copy warning, File not copied: {path_to}[/]')

    @staticmethod
    def last_modified_file(dir_path):
        files = [file for file in os.listdir(dir_path) if file != '.DS_Store']
        files = [os.path.join(dir_path, file) for file in files if os.path.isfile(os.path.join(dir_path, file))]
        return max(files, key=os.path.getctime) if files else print('[bold red]last modified file not found')

    @staticmethod
    def move(path_from, path_to):
        if os.path.exists(path_from):
            shutil.move(path_from, path_to)
            return print("[bold red]Move warning. File not moved") if not os.path.exists(path_to) else None
        return print(f"[bold red]Move warning. File not exist: {path_from}")

    @staticmethod
    def create_dir(path_to_dir, silence=False):
        if not os.path.exists(path_to_dir):
            os.makedirs(path_to_dir)
            if os.path.isdir(path_to_dir):
                return print(f'[green]Folder Created:[/] {path_to_dir}') if not silence else None
            return print(f'[bold red]Create folder warning. Folder not created: {path_to_dir}')

    @staticmethod
    def unpacking_via_7zip(archive_path, execute_path, delete=False):
        with py7zr.SevenZipFile(archive_path, 'r') as archive:
            archive.extractall(path=execute_path)
            print(f'[green]Unpack Completed.[/]')
        FileUtils.delete(archive_path, silence=True) if delete else ''

    @staticmethod
    def unpacking_via_zip_file(archive_path, execute_path, delete_archive=False):
        zip_archive = zipfile.ZipFile(archive_path)
        zip_archive.extractall(execute_path)
        zip_archive.close()
        FileUtils.delete(archive_path) if delete_archive else None

    @staticmethod
    def delete(what_delete, all_from_folder=False, silence=False):
        if not os.path.exists(what_delete):
            return print(f"[red]Delete warning. path not exist: {what_delete}") if not silence else None
        if os.path.isdir(what_delete):
            shutil.rmtree(what_delete, ignore_errors=True)
            if all_from_folder:
                FileUtils.create_dir(what_delete)
                if any(os.scandir(what_delete)):
                    return print(f"[red]Delete warning. Error while delete all from folder: {what_delete}")
                return print(f'[green]Folder is cleared') if not silence else None
        elif os.path.isfile(what_delete):
            os.remove(what_delete)
        if os.path.exists(what_delete):
            return print(f"[red]Delete warning. Folder exist: {what_delete}")
        print(f'[green]Object Deleted:[/] {what_delete}') if not silence else None

    @staticmethod
    def random_name(path, file_extension=None):
        while True:
            if file_extension:
                random_file_name = f'{random.randint(5000, 50000000)}.{file_extension}'
            else:
                random_file_name = f'{random.randint(5000, 50000000)}'
            if not os.path.exists(f'{path}/{random_file_name}'):
                return os.path.join(path, random_file_name)

    @staticmethod
    def run_command(command):
        popen = sb.Popen(command, stdout=sb.PIPE, stderr=sb.PIPE, shell=True)
        stdout, stderr = popen.communicate()
        popen.wait()
        stdout = stdout.strip().decode('utf-8', errors='ignore')
        stderr = stderr.strip().decode('utf-8', errors='ignore')
        popen.stdout.close(), popen.stderr.close()
        return stdout, stderr
