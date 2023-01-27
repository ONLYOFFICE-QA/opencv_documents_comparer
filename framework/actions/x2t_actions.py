# -*- coding: utf-8 -*-
import subprocess as sb
from os import chdir
from os.path import join

from configurations.project_configurator import ProjectConfig


class X2t:
    @staticmethod
    def x2t_version():
        chdir(ProjectConfig.core_dir())
        for line in sb.getoutput(join(ProjectConfig.core_dir(), 'x2t')).split("\n"):
            if ':' in line:
                key, value = line.strip().split(':', 1)
                if key.lower() == 'oox/binary file converter. version':
                    return value.strip()
        return ''
