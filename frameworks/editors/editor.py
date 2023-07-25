# -*- coding: utf-8 -*-
from abc import ABC, abstractmethod


class Editor(ABC):

    @property
    @abstractmethod
    def path(self) -> str: ...

    @property
    @abstractmethod
    def delay_after_open(self) -> int | float: ...

    @property
    @abstractmethod
    def process_names(self) -> list: ...

    @property
    @abstractmethod
    def formats(self) -> tuple | str: ...

    @abstractmethod
    def events_handler(self) -> object: ...

    @abstractmethod
    def open(self, file_path) -> None: ...

    @abstractmethod
    def make_screenshots(self, hwnd: int, screen_path: str, page_amount: int | dict) -> None: ...

    @abstractmethod
    def page_amount(self, file_path: str) -> int: ...

    @abstractmethod
    def close(self, hwnd: int) -> bool: ...

    @abstractmethod
    def set_size(self, hwnd: int) -> None: ...
