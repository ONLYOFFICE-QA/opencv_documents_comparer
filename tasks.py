# -*- coding: utf-8 -*-
from invoke import task
from rich import print
from rich.prompt import Prompt

import config
from frameworks.StaticData import StaticData

from host_control import HostInfo, File
from frameworks.editors.onlyoffice import Core, X2t
from frameworks.telegram import Telegram
from tests import X2tTesterConversion

if HostInfo().os == 'windows':
    from tests import CompareTest, OpenTests


@task
def download_core(c, force=False, version=None):
    version = version if version else config.version if config.version else Prompt.ask("Please enter version")
    Core(version).getting(force=force)


@task
def conversion_test(c, direction=None, ls=False, telegram=False, version=None, t_format=False, env_off=False):
    download_core(c, version=version)
    x2t_version = X2t.version(StaticData.core_dir())
    conversion = X2tTesterConversion(direction, x2t_version, trough_conversion=t_format, env_off=env_off)
    print(f"[bold green]|INFO| The conversion is running on x2t version: [red]{x2t_version}")
    report = conversion.from_files_list(config.files_array) if ls else conversion.run()
    tg_msg = f"Conversion completed on version: `{x2t_version}`\nPlatform: `{HostInfo().os}`" if telegram else None
    conversion.report.handler(report, x2t_version, tg_msg=tg_msg) if report else print("[red] Report not exists")
    print(f"[bold red]\n{'-' * 90}\n|INFO| x2t version: {x2t_version}\n{'-' * 90}")


@task
def make_files(c, telegram=False, direction=None, version=None):
    download_core(c, version=version)
    x2t_version = X2t.version(StaticData.core_dir())
    print(f"[bold green]|INFO| The files will be converted to x2t versions: [red]{x2t_version}")
    conversion = X2tTesterConversion(direction, x2t_version)
    report = conversion.run(results_path=True) if direction else conversion.from_extension_json()
    tg_msg = f"Files are converted to versions: `{x2t_version}`\nPlatform: `{HostInfo().os}`" if telegram else None
    conversion.report.handler(report, x2t_version, tg_msg=tg_msg) if report else print("[red] Report not exists")
    print(f"[bold red]\n{'-' * 90}\n|INFO| x2t version: {x2t_version}\n{'-' * 90}")


@task
def compare_test(c, direction: str = None, ls=False, telegram=False):
    direction = direction if direction else Prompt.ask('Input formats with -', default=None, show_default=False)
    source_ext, converted_ext = CompareTest().getting_formats(direction)

    if not source_ext or not converted_ext:
        raise print('[bold red]|ERROR| The direction is not correct')

    print("[bold green]|INFO| Starting...")
    CompareTest().run(
        File.get_paths(
            path=config.converted_docs,
            extension=converted_ext,
            dir_include=f"{config.version}_{source_ext}_",
            names=config.files_array if ls else None,
            exceptions_files=StaticData.ignore_files,
            exceptions_dirs=StaticData.ignore_dirs
        ),
        source_ext,
        converted_ext
    )

    if telegram:
        Telegram().send_message(f"Comparison on version {config.version} completed")


@task
def open_test(c, direction=None, ls=False, path='', telegram=False, new_test=False, fast_test=False):
    opener = OpenTests(continue_test=False if fast_test or path or new_test or ls else True)
    source_ext, converted_ext = opener.getting_formats(direction)

    if new_test:
        warning = f"[red]{'-' * 90}\nThe results will be removed from the report:\n{opener.report.path}\n{'-' * 90}\n"
        if Prompt.ask(warning, choices=['yes', 'no'], default='no') == 'yes':
            File.delete(opener.report.path)

    print("[bold green]|INFO| Starting...")
    opener.run(
        File.get_paths(
            path=path if path else config.converted_docs,
            extension=converted_ext if converted_ext else None,
            dir_include=f"{config.version}_{source_ext}_" if source_ext else config.version if not path else None,
            names=config.files_array if ls else None,
            exceptions_files=StaticData.ignore_files,
            exceptions_dirs=StaticData.ignore_dirs
        ),
        tg_msg=f"Opening test completed on version: {config.version}" if telegram else False
    )
