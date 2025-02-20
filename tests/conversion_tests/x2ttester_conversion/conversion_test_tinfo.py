# -*- coding: utf-8 -*-
from os import environ

from host_tools import HostInfo
from rich import print

from .x2ttester_test_config import X2ttesterTestConfig


class ConversionTestInfo:

    def __init__(self, test_config: X2ttesterTestConfig):
        self.config = test_config
        self.x2t_version = self.config.x2t_version
        self.quick_check = self.config.quick_check
        self.env_off = self.config.environment_off
        self.line = '-' * 90

    def out_test_info(self, mode: str):
        print(
            f"[bold green]|INFO| The conversion is running on x2t version: [red]{self.x2t_version}[/]\n"
            f"|INFO| Mode: [cyan]{mode}[/]\n"
            f"|INFO| X2t memory limit: [cyan]{environ.get('X2T_MEMORY_LIMIT', 'Default 4GIB')}[/]\n"
            f"|INFO| Environment: [cyan]{'True' if not self.env_off else 'False'}[/]"
        )

    def get_conversion_results_msg(self, version: str, execution_time: str, out: bool = True):
        msg =  (
        f"Conversion completed\n"
        f"Mode: `{'Quick Check' if self.quick_check else 'Full test'}`\n"
        f"Version: `{version}`\n"
        f"X2t version: `{self.x2t_version}`\n"
        f"Platform: `{HostInfo().os}`\n"
        f"Execution time: `{execution_time} min`"
        )

        if out:
            print(f"[green]{self.line}\n|INFO|{msg}\n{self.line}")

        return msg

    def get_make_files_result_msg(self, version: str, t_format: bool, out: bool = True):
        msg = (
            f"Files for open test converted\n"
            f"Version: `{version}`\n"
            f"X2t version: `{self.x2t_version}`\n"
            f"Platform: `{HostInfo().os}`\n"
            f"Mode: `{'t-format' if t_format else 'Default'}`"
        )

        if out:
            print(f"[green]{self.line}\n|INFO|{msg}\n{self.line}")

        return msg
