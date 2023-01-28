# -*- coding: utf-8 -*-

import codecs
import json
import random
import shutil
import subprocess as sb
import zipfile
from os import listdir, makedirs, scandir, remove, walk
from os.path import exists, isfile, isdir, join, getctime

import py7zr
from rich import print


class FileUtils:
    @staticmethod
    def read_json(path_to_json):
        with codecs.open(path_to_json, "r", "utf_8_sig") as file:
            return json.load(file)

    @staticmethod
    def delete_last_slash(path):
        return path.rstrip(path[-1]) if path[-1] in ['/', '\\'] else path

    @staticmethod
    def get_file_paths(dir_path):
        files_paths = []
        for root, dirs, files in walk(dir_path):
            for filename in files:
                files_paths.append(join(root, filename))
        return files_paths

    @staticmethod
    def get_files_by_extensions(path_to_folder, ext):
        files_paths = []
        for root, dirs, files in walk(path_to_folder):
            for filename in files:
                if ext and filename.lower().endswith(ext if isinstance(ext, tuple) else ext.lower()):
                    files_paths.append(join(root, filename))
        return files_paths

    # path insert with file name
    @staticmethod
    def copy(path_from, path_to, silence=False):
        if not exists(path_from):
            return print(f"[red]copy warning: path not exist: {path_from}")
        if isdir(path_from):
            shutil.copytree(path_from, path_to)
        elif isfile(path_from):
            shutil.copyfile(path_from, path_to)
        if exists(path_to):
            return print(f'[green]Copied To: {path_to}') if not silence else None
        return print(f'[red]copy warning, File not copied: {path_to}')

    @staticmethod
    def last_modified_file(dir_path):
        files = [file for file in listdir(dir_path) if file != '.DS_Store']
        files = [join(dir_path, file) for file in files if isfile(join(dir_path, file))]
        return max(files, key=getctime) if files else print('[bold red]last modified file not found')

    @staticmethod
    def move(path_from, path_to):
        if exists(path_from):
            shutil.move(path_from, path_to)
            return print("[bold red]Move warning. File not moved") if not exists(path_to) else None
        return print(f"[bold red]Move warning. File not exist: {path_from}")

    @staticmethod
    def create_dir(path_to_dir, silence=False):
        if not exists(path_to_dir):
            makedirs(path_to_dir)
            if isdir(path_to_dir):
                return print(f'[green]Folder Created: {path_to_dir}') if not silence else None
            return print(f'[bold red]Create folder warning. Folder not created: {path_to_dir}')

    @staticmethod
    def unpacking_via_7zip(archive_path, execute_path, delete=False):
        with py7zr.SevenZipFile(archive_path, 'r') as archive:
            archive.extractall(path=execute_path)
            print(f'[green]Unpack Completed.')
        FileUtils.delete(archive_path, silence=True) if delete else ''

    @staticmethod
    def unpacking_via_zip_file(archive_path, execute_path, delete_archive=False):
        zip_archive = zipfile.ZipFile(archive_path)
        zip_archive.extractall(execute_path)
        zip_archive.close()
        FileUtils.delete(archive_path) if delete_archive else None

    @staticmethod
    def delete(what_delete, all_from_folder=False, silence=False):
        if not exists(what_delete):
            return print(f"[red]Delete warning. path not exist: {what_delete}") if not silence else None
        if isdir(what_delete):
            shutil.rmtree(what_delete, ignore_errors=True)
            if all_from_folder:
                FileUtils.create_dir(what_delete)
                if any(scandir(what_delete)):
                    return print(f"[red]Delete warning. Error while delete all from folder: {what_delete}")
                return print(f'[green]Folder is cleared') if not silence else None
        elif isfile(what_delete):
            remove(what_delete)
        if exists(what_delete):
            return print(f"[red]Delete warning. Folder exist: {what_delete}")
        print(f'[green]Deleted: {what_delete}') if not silence else None

    @staticmethod
    def random_name(path, extension=None):
        while True:
            random_name = f'{random.randint(500, 50000)}.{extension}' if extension else f'{random.randint(500, 50000)}'
            random_object_path = join(path, random_name)
            if not exists(random_object_path):
                return random_object_path

    @staticmethod
    def run_command(command):
        popen = sb.Popen(command, stdout=sb.PIPE, stderr=sb.PIPE, shell=True)
        stdout, stderr = popen.communicate()
        popen.wait()
        stdout = stdout.strip().decode('utf-8', errors='ignore')
        stderr = stderr.strip().decode('utf-8', errors='ignore')
        popen.stdout.close(), popen.stderr.close()
        return stdout, stderr
