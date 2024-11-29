# -*- coding: utf-8 -*-
from typing import Optional
from pydantic import BaseModel

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
