# -*- coding: utf-8 -*-
from host_tools import HostInfo
from tests.conversion_tests.x2ttester_conversion import X2tTesterConversion

if HostInfo().os == 'windows':
    from .compare_tests import CompareTest
    from .open_tests import OpenTests
