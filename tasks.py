# -*- coding: utf-8 -*-
import time
from os import environ

from invoke import task
from rich import print
from rich.prompt import Prompt

import config
from frameworks.StaticData import StaticData

from host_tools import HostInfo, File
from frameworks.editors.onlyoffice import Core, X2t
from telegram import Telegram

from frameworks.s3 import S3Downloader
from tests import X2tTesterConversion

if HostInfo().os == 'windows':
    from tests import CompareTest, OpenTests


@task
def download_core(c, force=False, version=None):
    version = version if version else config.version if config.version else Prompt.ask("Please enter version")
    Core(version).getting(force=force)


@task
def download_files(c, cores: int = None, sha256: bool = False):
    S3Downloader(download_dir=config.source_docs, cores=cores, check_sha256=sha256).download_all()


@task
def conversion_test(
        c,
        direction: str = None,
        ls: bool = False,
        telegram: bool = False,
        version: str = None,
        t_format: bool = False,
        env_off: bool = False,
        quick_check: bool = False,
        x2t_limits: int = None
):
    if x2t_limits and not env_off:
        environ['X2T_MEMORY_LIMIT'] = f"{x2t_limits}GiB"

    download_core(c, version=version)

    x2t_version = X2t.version(StaticData.core_dir())
    print(
        f"[bold green]|INFO| The conversion is running on x2t version: [red]{x2t_version}[/]\n"
        f"|INFO| Mode: "
        f"{'[cyan]Quick Check' if quick_check else '[red]Full test' if not ls else '[magenta]From array'}[/]\n"
        f"|INFO| X2t memory limit: [cyan]{environ.get('X2T_MEMORY_LIMIT', 'Default 4GIB')}[/]\n"
        f"|INFO| Environment: [cyan]{'True' if not env_off else 'False'}[/]"
    )

    conversion = X2tTesterConversion(direction, x2t_version, trough_conversion=t_format, env_off=env_off)
    files_list = conversion.get_quick_check_files() if quick_check else config.files_array if ls else None
    object_keys = [f"{name.split('.')[-1].lower()}/{name}" for name in files_list] if files_list else None
    S3Downloader(download_dir=config.source_docs).download_all(objects=object_keys)

    start_time = time.perf_counter()
    report = conversion.from_files_list(files_list) if files_list else conversion.run()

    results_msg = (
        f"Conversion completed\n"
        f"Mode: `{'Quick Check' if quick_check else 'Full test'}`"
        f"Version: `{x2t_version}`\n"
        f"Platform: `{HostInfo().os}`\n"
        f"Execution time: `{((time.perf_counter() - start_time) / 60):.02f} min`"
    )

    if report:
        conversion.report.handler(report, x2t_version, tg_msg=results_msg if telegram else None)

    print(f"[green]{'-' * 90}\n|INFO|{results_msg}\n{'-' * 90}")


@task
def make_files(c, telegram=False, direction=None, version=None, t_format=False, env_off=False):
    download_core(c, version=version)

    x2t_version = X2t.version(StaticData.core_dir())
    print(f"[bold green]|INFO| The files will be converted to x2t versions: [red]{x2t_version}")

    conversion = X2tTesterConversion(direction, x2t_version, trough_conversion=t_format, env_off=env_off)
    report = conversion.run(results_path=True) if direction else conversion.from_extension_json()

    tg_msg = (
        f"Files for open test converted\n"
        f"Version: `{x2t_version}`\n"
        f"Platform: `{HostInfo().os}`\n"
        f"Mode: `{'t-format' if t_format else 'Default'}`"
    ) if telegram else None

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
            dir_include=f"{config.version}_(dir_{source_ext}",
            names=config.files_array if ls else None,
            exceptions_files=StaticData.ignore_files,
            exceptions_dirs=StaticData.ignore_dirs
        ),
        source_ext,
        converted_ext
    )

    if telegram:
        Telegram().send_message(f"Comparison completed\nVersion: {config.version}")


@task
def open_test(c, version=None, direction=None, ls=False, path='', telegram=False, new_test=False, fast_test=False):
    _version = version if version else config.version
    opener = OpenTests(_version, continue_test=False if fast_test or path or new_test or ls else True)
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
            dir_include=f"{config.version}_(dir_{source_ext}" if source_ext else config.version if not path else None,
            names=config.files_array if ls else None,
            exceptions_files=StaticData.ignore_files,
            exceptions_dirs=StaticData.ignore_dirs
        ),
        tg_msg=f"Opening test completed\nVersion: {config.version}" if telegram else False
    )
