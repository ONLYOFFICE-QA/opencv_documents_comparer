# -*- coding: utf-8 -*-

import os
import json
import shutil
import subprocess as sb
import codecs
import pyautogui as pg
from rich import print


class FileUtils:
    @staticmethod
    def read_json(path_to_json):
        with codecs.open(path_to_json, "r", "utf_8_sig") as file:
            return json.load(file)

    @staticmethod
    def click(path_to_image):
        from data.StaticData import StaticData
        try:
            pg.click(f'{StaticData.PROJECT_FOLDER}/data/image_templates/{path_to_image}')
            return True
        except TypeError:
            return False

    # path insert with file name
    @staticmethod
    def copy(path_from, path_to):
        if os.path.exists(path_from):
            shutil.copyfile(path_from, path_to)
            return print("[bold red]Copy warning. File not copied") if not os.path.exists(path_to) else None
        return print(f"[bold red]Copy warning. File not exist: {path_from}")

    @staticmethod
    def move(path_from, path_to):
        if os.path.exists(path_from):
            shutil.move(path_from, path_to)
            return print("[bold red]Move warning. File not moved") if not os.path.exists(path_to) else None
        return print(f"[bold red]Move warning. File not exist: {path_from}")

    @staticmethod
    def create_dir(path_to_dir):
        if not os.path.exists(path_to_dir):
            os.makedirs(path_to_dir)
            if not os.path.isdir(path_to_dir):
                print(f'[bold red]Crate folder warning. Folder not created: {path_to_dir}')

    @staticmethod
    def delete(what_delete, all_from_folder=False):
        if os.path.isdir(what_delete):
            shutil.rmtree(what_delete, ignore_errors=True)
            if all_from_folder:
                FileUtils.create_dir(what_delete)
                if any(os.scandir(what_delete)):
                    print(f"[bold red]delete warning. Error while delete all from folder. Not empty: {what_delete}")
        elif os.path.isfile(what_delete):
            os.remove(what_delete)
        print(f"[bold red]Error while delete: {what_delete}") if os.path.exists(what_delete) else None

    @staticmethod
    def run_command(command):
        popen = sb.Popen(command, stdout=sb.PIPE, stderr=sb.PIPE, shell=True)
        stdout, stderr = popen.communicate()
        popen.wait()
        stdout = stdout.strip().decode('utf-8', errors='ignore')
        stderr = stderr.strip().decode('utf-8', errors='ignore')
        popen.stdout.close(), popen.stderr.close()
        return stdout, stderr
