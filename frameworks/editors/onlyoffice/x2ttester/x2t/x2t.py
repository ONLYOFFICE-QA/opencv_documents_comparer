# -*- coding: utf-8 -*-
from os import chdir
from os.path import isfile, join
from typing import Optional
from host_tools import Str, Shell
from ..host_config import HostConfig


class X2t:
    @staticmethod
    def version(x2t_dir: str) -> Optional[str]:
        chdir(x2t_dir)
        x2t_info = Shell.get_output(X2t.check_exists(join(x2t_dir, HostConfig().x2t)))
        return Str.find_by_key(x2t_info, key='OOX/binary file converter. Version')

    @staticmethod
    def check_exists(x2t_path: str) -> Optional[str]:
        if isfile(x2t_path):
            return x2t_path
        raise FileNotFoundError(f"[bold red]|WARNING| X2t not exists, check path: {x2t_path}")
