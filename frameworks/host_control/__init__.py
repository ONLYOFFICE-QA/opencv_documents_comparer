# -*- coding: utf-8 -*-
from .utils import FileUtils
from .host_info import HostInfo

if HostInfo().os == 'windows':
    from .windows_handler import Window
