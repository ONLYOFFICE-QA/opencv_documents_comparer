# -*- coding: utf-8 -*-
from host_control import HostInfo
from .x2ttester_conversion import X2tTesterConversion

if HostInfo().os == 'windows':
    from .compare_tests import CompareTest
    from .open_tests import OpenTests
