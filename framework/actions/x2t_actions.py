# -*- coding: utf-8 -*-
from os import chdir
from os.path import exists
from os.path import join

from framework.StaticData import StaticData
from framework.FileUtils import FileUtils
from framework.actions.host_actions import HostActions


class X2t:
    @staticmethod
    def x2t_version(x2t_dir=StaticData.core_dir()):
        host = HostActions()
        if exists(join(x2t_dir, host.x2t)):
            chdir(x2t_dir)
            x2t_info = FileUtils.output_cmd(join(StaticData.core_dir(), host.x2t))
            return FileUtils.find_in_line_by_key(x2t_info, 'OOX/binary file converter. Version')
