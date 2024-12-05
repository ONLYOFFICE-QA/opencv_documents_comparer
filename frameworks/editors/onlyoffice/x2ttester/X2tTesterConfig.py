# -*- coding: utf-8 -*-
from typing import Optional
from pydantic import BaseModel, model_validator
from rich import print

from frameworks.decorators import highlighter


class X2tTesterConfig(BaseModel):
    input_dir: str
    output_dir: str
    core_dir: str
    report_path: str
    fonts_dir: Optional[str] = None
    cores: int = 1
    timeout: Optional[int] = None
    timestamp: bool = True
    delete: bool = False
    errors_only: bool = False
    trough_conversion: bool = False
    environment_off: bool = False
    out_x2ttester_param: bool = False

    @model_validator(mode="after")
    def set_computed_fields(self):
        if self.out_x2ttester_param:
            self._out_param()
        return self

    @highlighter()
    def _out_param(self):
        print("[magenta]|INFO| X2ttester Parameters:")
        for field, value in self.model_dump().items():
            if field != 'out_x2ttester_param':
                print(f"[green]{field}: [cyan]{value}[/]")
