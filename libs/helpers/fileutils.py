# -*- coding: utf-8 -*-
from management import *


class FileUtils:
    @staticmethod
    def read_json(path_to_json):
        with codecs.open(path_to_json, "r", "utf_8_sig") as file:
            json_data = json.load(file)
        return json_data

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
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.copyfile(path_from, path_to)

    @staticmethod
    def move(path_from, path_to):
        if os.path.exists(path_from) and not os.path.exists(path_to):
            shutil.move(path_from, path_to)

    @staticmethod
    def create_dir(path_to_dir):
        if not os.path.exists(path_to_dir):
            os.makedirs(path_to_dir)

    @staticmethod
    def delete(what_delete):
        sb.call(f'powershell.exe rm {what_delete} -Force -Recurse', shell=True)

    @staticmethod
    def dict_compare(source_statistics, converted_statistics):
        d1_keys, d2_keys = set(source_statistics.keys()), set(converted_statistics.keys())
        modified = {
            o: (f'Source {source_statistics[o]}', f'Converted {converted_statistics[o]}')
            for o in d1_keys.intersection(d2_keys) if source_statistics[o] != converted_statistics[o]
        }
        return modified

    @staticmethod
    def run_command(open_command):
        sb.Popen(open_command)
