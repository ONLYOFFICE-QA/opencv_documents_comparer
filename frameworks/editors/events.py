# -*- coding: utf-8 -*-
from abc import ABC, abstractmethod


class Events(ABC):

    @property
    @abstractmethod
    def window_class_names(self) -> list: ...

    @abstractmethod
    def when_opening(self, class_name: str, windows_text: str, hwnd: int = None) -> bool: ...

    @abstractmethod
    def when_closing(self, class_name, windows_text, hwnd: int = None) -> bool: ...
