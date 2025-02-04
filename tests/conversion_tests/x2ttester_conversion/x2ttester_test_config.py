# -*- coding: utf-8 -*-
from os import environ, getenv
from os.path import dirname, join, splitdrive
from tempfile import gettempdir

from host_tools import File, HostInfo
from host_tools.utils import Dir
from pydantic import BaseModel, Field, model_validator
from typing import Optional

import config
from frameworks.StaticData import StaticData
from frameworks.editors.onlyoffice import VersionHandler, X2t



class X2ttesterTestConfig(BaseModel):
    core_dir: Optional[str] = None
    reports_dir: Optional[str] = None
    input_dir: Optional[str] = None
    output_dir: Optional[str] = Field(init=False, default=None)
    result_dir: Optional[str] = None
    cores: Optional[int] = None
    timeout: Optional[int] = None
    delete: Optional[bool] = None
    direction: Optional[str] = Field(default=None, description="Conversion direction, e.g., 'docx-pdf'")
    trough_conversion: bool = False
    environment_off: bool = False
    tmp_dir: Optional[str] = None
    report_path: Optional[str] = Field(init=False, default=None)
    timestamp: bool = Field(init=False, default=False)
    fonts_dir: Optional[str] = Field(init=False, default=None)
    errors_only: bool = Field(init=False, default=True)
    x2t_version: Optional[str] = Field(init=False, default=None)
    output_formats: Optional[str] = Field(init=False, default=None)
    input_formats: Optional[str] = Field(init=False, default=None)
    out_x2ttester_param: bool = False
    x2t_memory_limits: Optional[int] = None
    quick_check: bool = False

    @model_validator(mode="after")
    def set_computed_fields(self):
        self.core_dir = self.core_dir or StaticData.core_dir()
        self.reports_dir = self.reports_dir or StaticData.reports_dir()
        self.x2t_version = VersionHandler(X2t.version(self.core_dir)).version
        self.tmp_dir = File.unique_name(self.tmp_dir or self._get_tmp_dir())
        self.report_path = self.get_tmp_report_path()
        self.cores = self.cores or config.cores
        self.timeout = self.timeout or config.timeout
        self.timestamp = config.timestamp
        self.delete = config.delete if self.delete is None else self.delete
        self.errors_only = config.errors_only
        self.output_dir = join(self.tmp_dir, 'cnv')
        self.result_dir = self.result_dir or StaticData.result_dir()
        self.input_dir = self.input_dir or StaticData.documents_dir()
        self.fonts_dir = StaticData.fonts_dir()
        self.input_formats, self.output_formats = self._getting_formats(self.direction)
        self._set_x2t_memory_limits()
        return self

    def get_tmp_report_path(self) -> str:
        tmp_report = File.unique_name(File.unique_name(self.tmp_dir), 'csv')
        Dir.create(dirname(tmp_report), stdout=False)
        return tmp_report

    @staticmethod
    def _getting_formats(direction: str | None = None) -> tuple[None | str, None | str]:
        if direction:
            if '-' in direction:
                return direction.split('-')[0], direction.split('-')[1]
            return None, direction
        return None, None

    def _set_x2t_memory_limits(self):
        if self.x2t_memory_limits and not self.environment_off:
            environ['X2T_MEMORY_LIMIT'] = f"{self.x2t_memory_limits}GiB"

    @staticmethod
    def _get_tmp_dir() -> str:
        return join(f"{splitdrive(getenv('SystemRoot'))[0]}/", "tmp") if HostInfo().os == 'windows' else gettempdir()
