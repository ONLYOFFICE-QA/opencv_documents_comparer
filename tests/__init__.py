# -*- coding: utf-8 -*-
from host_tools import HostInfo
from .conversion_tests import X2tTesterConversion, X2ttesterTestConfig, ErrorHandler

__all__ = ['X2tTesterConversion', 'X2ttesterTestConfig', 'ErrorHandler']

if HostInfo().os == 'windows':
    from .compare_tests import CompareTest
    from .open_tests import OpenTests
    __all__.extend(['CompareTest', 'OpenTests'])
