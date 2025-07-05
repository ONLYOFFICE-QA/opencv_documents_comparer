# -*- coding: utf-8 -*-
import time
from typing import Optional

from invoke import task
from rich import print
from rich.prompt import Prompt

import config
from frameworks.StaticData import StaticData

from host_tools import HostInfo, File
from frameworks.editors.onlyoffice import Core, X2t
from telegram import Telegram

from frameworks.s3 import S3Downloader, S3Uploader
from tests import X2tTesterConversion, X2ttesterTestConfig

if HostInfo().os == 'windows':
    from tests import CompareTest, OpenTests


@task
def download_core(c, force: bool = False, version: Optional[str] = None):
    Core(version=version or Prompt.ask("Please enter version")).getting(force=force)


@task
def download_files(c, cores: Optional[int] = None, sha256: bool = False):
    S3Downloader(download_dir=config.source_docs, cores=cores, check_sha256=sha256).download_all()


@task
def conversion_test(
        c,
        cores: Optional[int] = None,
        direction: Optional[str] = None,
        ls: bool = False,
        telegram: bool = False,
        version: Optional[str] = None,
        t_format: bool = False,
        env_off: bool = False,
        quick_check: bool = False,
        x2t_limits: Optional[int] = None,
        out_x2ttester_param: bool = False
):
    download_core(c, version=version)
    conversion = X2tTesterConversion(
        test_config=X2ttesterTestConfig(
            cores=cores,
            delete=True,
            direction=direction,
            environment_off=env_off,
            trough_conversion=t_format,
            out_x2ttester_param=out_x2ttester_param,
            x2t_memory_limits=x2t_limits,
            quick_check=quick_check
        )
    )

    conversion.info.out_test_info(mode='Quick Check' if quick_check else 'From array' if  ls else 'Full test')

    files_list = conversion.get_quick_check_files() if quick_check else config.files_array if ls else None
    object_keys = [f"{name.split('.')[-1].lower()}/{name}" for name in files_list] if files_list else None
    S3Downloader(download_dir=conversion.config.input_dir).download_all(objects=object_keys)

    start_time = time.perf_counter()
    report = conversion.from_files_list(files_list) if files_list else conversion.run()
    execution_time = f"{((time.perf_counter() - start_time) / 60):.02f}"

    results_msg = conversion.info.get_conversion_results_msg(str(version), execution_time)

    if report:
        conversion.report.handler(report_path=report, tg_msg=results_msg if telegram else None)


@task
def make_files(
        c,
        cores: int = None,
        telegram: bool = False,
        direction: str = None,
        version: str = None,
        t_format: bool = False,
        env_off: bool = False,
        full: bool = False,
        out_x2ttester_param: bool = False
):
    download_core(c, version=version)
    conversion = X2tTesterConversion(
        test_config=X2ttesterTestConfig(
            cores=cores,
            delete=False,
            direction=direction,
            environment_off=env_off,
            trough_conversion=t_format,
            out_x2ttester_param=out_x2ttester_param
        )
    )

    conversion.info.out_test_info(mode='Make files for openers')

    S3Downloader(download_dir=conversion.config.input_dir).download_all()

    report = conversion.run(results_path=True) if direction else conversion.from_extension_json()

    if full and not t_format:
        conversion.config.trough_conversion = True
        conversion = X2tTesterConversion(conversion.config)
        report = conversion.run(results_path=True) if direction else conversion.from_extension_json()

    tg_msg = conversion.info.get_make_files_result_msg(version=version, t_format=t_format) if telegram else None
    conversion.report.handler(report, tg_msg=tg_msg) if report else print("[red] Report not exists")
    print(f"[bold red]\n{'-' * 90}\n|INFO| x2t version: {X2t.version(conversion.config.core_dir)}\n{'-' * 90}")


@task
def compare_test(c, direction: str = None, ls: bool = False, telegram: bool = False):
    direction = direction or Prompt.ask('Input formats with -', default=None, show_default=False)
    source_ext, converted_ext = CompareTest.getting_formats(direction)

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
def open_test(
        c,
        version: str = None,
        direction: str = None,
        ls: bool = False,
        path: str = None,
        telegram: bool = False,
        new_test:  bool = False,
        fast_test: bool = False
):
    version = version or config.version or Prompt.ask("Please enter version")

    opener = OpenTests(version, continue_test=False if fast_test or path or new_test or ls else True)
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
            dir_include=f"{version}_(dir_{source_ext}" if source_ext else version if not path else None,
            names=config.files_array if ls else None,
            exceptions_files=StaticData.ignore_files,
            exceptions_dirs=StaticData.ignore_dirs
        ),
        tg_msg=f"Opening test completed\nVersion: {version}" if telegram else False
    )

@task
def s3_upload(c, dir_path: str, cores: int = None):
    uploader = S3Uploader(
        bucket_name='conversion-testing-files',
        region='us-east-1',
        check_duplicates=True,
        cores=cores
    )
    uploader.upload_from_dir(dir_path=dir_path)
