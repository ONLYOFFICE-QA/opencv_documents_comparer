# -*- coding: utf-8 -*-
import subprocess as sb
from os import chdir
from os.path import exists
from os.path import join

from configurations.project_configurator import ProjectConfig
from framework.actions.host_actions import HostActions


class X2t:
    @staticmethod
    def x2t_version():
        host = HostActions()
        if exists(join(ProjectConfig.core_dir(), host.x2t)):
            chdir(ProjectConfig.core_dir()) if host.os == 'mac' else None
            for line in sb.getoutput(join(ProjectConfig.core_dir(), host.x2t)).split("\n"):
                if ':' in line:
                    key, value = line.strip().split(':', 1)
                    if key.lower() == 'oox/binary file converter. version':
                        return value.strip()
        return ''
