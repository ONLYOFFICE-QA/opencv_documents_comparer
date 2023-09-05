# -*- coding: utf-8 -*-
import platform
import distro

from ..singleton import singleton


@singleton
class Unix:
    def __init__(self):
        self._name = distro.name()
        self._version = distro.version()
        self._id = distro.id()
        self._pretty_name = distro.name(pretty=True)

    @property
    def pretty_name(self) -> str:
        return self._pretty_name

    @pretty_name.setter
    def pretty_name(self, value: str) -> None:
        if value:
            self._pretty_name = value
        else:
            raise print(f"[bold red]|ERROR| Can't get pretty name")

    @property
    def id(self) -> str:
        return self._id

    @id.setter
    def id(self, value: str):
        if value:
            self._id = value
        else:
            raise print(f"[bold red]|ERROR| Can't get distribution id")

    @property
    def name(self) -> str:
        return self._name

    @property
    def version(self) -> str:
        return self._version

    @version.setter
    def version(self, value: str):
        if value:
            self._version = value
        else:
            raise print(f"[bold red]|ERROR| Can't get distribution version")

    @name.setter
    def name(self, value: str):
        if value:
            self._name = value
        else:
            raise print(f"[bold red]|ERROR| Can't get distribution name")

    @staticmethod
    def get_distro_info() -> dict:
        return platform.freedesktop_os_release()
