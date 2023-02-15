# -*- coding: utf-8 -*-
import codecs
import json
import string
import zipfile
from os import listdir, makedirs, scandir, remove, walk
from os.path import exists, isfile, isdir, join, getctime, basename, getsize, relpath
from random import randint, choice
from shutil import move, copytree, copyfile, rmtree
from subprocess import Popen, PIPE, getoutput

import py7zr
from requests import get
from rich import print
from rich.progress import track


class FileUtils:
    @staticmethod
    def read_json(path_to_json):
        with codecs.open(path_to_json, "r", "utf_8_sig") as file:
            return json.load(file)

    @staticmethod
    def write_json(path_to, data, mode='w'):
        if not exists(path_to):
            return print(f"|WARNING|The path to json file does not exist: {path_to}")
        with open(path_to, mode) as file:
            json.dump(data, file, indent=2)

    @staticmethod
    def delete_last_slash(path):
        return path.rstrip(path[-1]) if path[-1] in ['/', '\\'] else path

    @staticmethod
    def get_file_paths(dir_path, ext=''):
        files_paths = []
        for root, dirs, files in walk(dir_path):
            for filename in files:
                if ext and filename.lower().endswith(ext if isinstance(ext, tuple) else ext.lower()):
                    files_paths.append(join(root, filename))
                elif not ext:
                    files_paths.append(join(root, filename))
        return files_paths

    # path insert with file name
    @staticmethod
    def copy(path_from, path_to, silence=False):
        if not exists(path_from):
            return print(f"[bold red]|COPY WARNING| Path from not exist: {path_from}")
        if isdir(path_from):
            copytree(path_from, path_to)
        elif isfile(path_from):
            copyfile(path_from, path_to)
        if exists(path_to):
            return print(f'[green]|INFO| Copied to: {path_to}') if not silence else ...
        return print(f'[bold red]|COPY WARNING| File not copied: {path_to}')

    @staticmethod
    def last_modified_file(dir_path):
        files = [file for file in listdir(dir_path) if file != '.DS_Store']
        files = [join(dir_path, file) for file in files if isfile(join(dir_path, file))]
        return max(files, key=getctime) if files else print('[bold red]|WARNING| Last modified file not found')

    @staticmethod
    def move(path_from, path_to):
        if exists(path_from):
            move(path_from, path_to)
            return print("[bold red]|MOVE WARNING| File not moved") if not exists(path_to) else ...
        return print(f"[bold red]|MOVE WARNING| File not exist: {path_from}")

    @staticmethod
    def create_dir(path_to_dir, silence=False):
        if not exists(path_to_dir):
            makedirs(path_to_dir)
            if isdir(path_to_dir):
                return print(f'[green]|INFO| Folder Created: {path_to_dir}') if not silence else ...
            return print(f'[bold red]|WARNING| Create folder warning. Folder not created: {path_to_dir}')

    @staticmethod
    def unpacking_via_7zip(archive_path, execute_path, delete=False):
        print(f'[green]|INFO| Unpacking {basename(archive_path)}.')
        with py7zr.SevenZipFile(archive_path, 'r') as archive:
            archive.extractall(path=execute_path)
            print(f'[green]|INFO| Unpack Completed.')
        FileUtils.delete(archive_path, silence=True) if delete else ...

    @staticmethod
    def compress_files(path, archive_path=None, delete=False):
        archive = archive_path if archive_path else f"{path}.zip"
        if not exists(path):
            return print(f'[bold red]|COMPRESS WARNING|Path for compression does not exist: {path}')
        print(f'[green]|INFO|Compressing: {path}')
        with zipfile.ZipFile(archive, 'w') as zip_archive:
            if isfile(path):
                zip_archive.write(path, basename(path), compress_type=zipfile.ZIP_DEFLATED)
            elif isdir(path):
                for file in track(FileUtils.get_file_paths(path), description=f"[red]Compressing dir:{basename(path)}"):
                    zip_archive.write(file, relpath(file, path), compress_type=zipfile.ZIP_DEFLATED)
            else:
                return print(f"|WARNING|The path for archiving is neither a file nor a directory. {path}")
        if exists(archive) and getsize(archive) != 0:
            FileUtils.delete(path) if delete else ...
            return print(f"[green]|INFO|Success compressed: {archive}")
        print(f"[WARNING]Archive not exists: {archive}")

    @staticmethod
    def unpacking_via_zip_file(archive_path, execute_path, delete_archive=False):
        with zipfile.ZipFile(archive_path) as zip_archive:
            zip_archive.extractall(execute_path)
        FileUtils.delete(archive_path) if delete_archive else ...

    @staticmethod
    def delete(what_delete, all_from_folder=False, silence=False):
        if not exists(what_delete):
            return print(f"[bold red]|DELETE WARNING| Path not exist: {what_delete}") if not silence else ...
        if isdir(what_delete):
            rmtree(what_delete, ignore_errors=True)
            if all_from_folder:
                FileUtils.create_dir(what_delete)
                if any(scandir(what_delete)):
                    return print(f"[bold red]|DELETE WARNING| Error while delete all from folder: {what_delete}")
                return print(f'[green]|INFO| Folder is cleared') if not silence else ...
        elif isfile(what_delete):
            remove(what_delete)
        if exists(what_delete):
            return print(f"[bold red]|DELETE WARNING| Folder is not deleted: {what_delete}")
        print(f'[green]|INFO| Deleted: {what_delete}') if not silence else ...

    @staticmethod
    def random_string(path_to_dir, chars=string.ascii_uppercase + string.digits, num_chars=50, extension=None):
        while True:
            random_string = ''.join(choice(chars).lower() for _ in range(int(num_chars)))
            random_object_path = join(path_to_dir, f"{random_string}.{extension}" if extension else random_string)
            if not exists(random_object_path):
                return random_object_path

    @staticmethod
    def random_name(path, extension=None):
        while True:
            random_name = f'{randint(500, 50000)}.{extension}' if extension else f'{randint(500, 50000)}'
            random_object_path = join(path, random_name)
            if not exists(random_object_path):
                return random_object_path

    @staticmethod
    def download_file(url, path, name=None):
        file_path, _ = join(path, name if name else basename(url)), FileUtils.create_dir(path)
        with get(url, stream=True) as r:
            r.raise_for_status()
            with open(file_path, 'wb') as file:
                for chunk in track(r.iter_content(chunk_size=1024 * 1024), description=f'[red]Downloading:{name}'):
                    if chunk:
                        file.write(chunk)
        print(f"[bold green]|INFO|File Saved to: {file_path}" if isfile(file_path) else f"[red]|WARNING|File not saved")
        if int(getsize(file_path)) != int(r.headers['Content-Length']):
            print(f"[red]|WARNING|Size different\nFile:{getsize(file_path)}\nOn server:{r.headers['Content-Length']}")

    @staticmethod
    def find_in_line_by_key(text, key_trigger, split_by='\n', separator=':'):
        for line in text.split(split_by):
            if separator in line:
                key, value = line.strip().split(separator, 1)
                if key.lower() == key_trigger.lower():
                    return value.strip()

    @staticmethod
    def file_reader(file_path, mode='r'):
        with open(file_path, mode) as file:
            return file.read()

    @staticmethod
    def file_writer(file_path, text, mode='w'):
        with open(file_path, mode) as file:
            file.write(text)

    @staticmethod
    def output_cmd(command):
        return getoutput(command)

    @staticmethod
    def run_command(command):
        popen = Popen(command, stdout=PIPE, stderr=PIPE, shell=True)
        stdout, stderr = popen.communicate()
        popen.wait()
        stdout = stdout.strip().decode('utf-8', errors='ignore')
        stderr = stderr.strip().decode('utf-8', errors='ignore')
        popen.stdout.close(), popen.stderr.close()
        return stdout, stderr
