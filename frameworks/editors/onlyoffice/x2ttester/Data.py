# -*- coding: utf-8 -*-
from pydantic import BaseModel


class Data(BaseModel):
    input_dir: str
    output_dir: str
    x2ttester_dir: str
    report_path: str
    fonts_dir: str = None
    cores: int = 1
    timeout: int = None
    timestamp: bool = True
    delete: bool = False
    errors_only: bool = False
    trough_conversion: bool = False
    save_environment: bool = False
