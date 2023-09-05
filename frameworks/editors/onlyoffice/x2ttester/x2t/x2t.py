# -*- coding: utf-8 -*-
from os import chdir
from os.path import exists
from os.path import join
from rich import print

from frameworks.StaticData import StaticData
from frameworks.host_control import FileUtils

from ..host_config import HostConfig


class X2t:
    @staticmethod
    def version(x2t_dir: str) -> str | None:
        if exists(join(x2t_dir, HostConfig().x2t)):
            chdir(x2t_dir)
            x2t_info = FileUtils.output_cmd(join(StaticData.core_dir(), HostConfig().x2t))
            return FileUtils.find_in_line_by_key(x2t_info, key='OOX/binary file converter. Version')
        print(f"[bold red]|WARNING| X2t not exists, check path: {join(x2t_dir, HostConfig().x2t)}")
