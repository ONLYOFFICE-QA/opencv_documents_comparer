# -*- coding: utf-8 -*-
from re import sub


class VersionHandler:
    """
    Class for handling version numbers like “00.00.00.00”.

    Provides functionality to parse version numbers and extract major, minor and build components.

    Attributes:
        _version_pattern (str): Regular expression pattern to match version numbers.
        Version (str): Version number string.
    """

    def __init__(self, version: str):
        """
        Initializes the VersionHandler object.
        :param version: Version number string.
        """
        self._version_pattern = r'(\d+).(\d+).(\d+).(\d+)'
        self.version = version

    @property
    def version(self) -> str:
        """
        Getter for the version number.
        :return: Version number string.
        """
        return self.__version

    @version.setter
    def version(self, value: str) -> None:
        """
        Setter for the version number.
        Validates the format of the version number.
        :param value: Version number string.
        :raises: ValueError if the version number is not in the correct format.
        """
        if len([i for i in value.split('.') if i]) != 4:
            raise ValueError(
                "[red]|WARNING| Version is entered incorrectly. The version must be in the format 'x.x.x.x'"
            )
        self.__version = value

    @property
    def major(self) -> str:
        """
        Extracts the major version component from the version number.
        :return: Major version string.
        """
        return sub(self._version_pattern, r'\1.\2', self.version)

    @property
    def minor(self) -> str:
        """
        Extracts the minor version component from the version number.
        :return: Minor version string.
        """
        return sub(self._version_pattern, r'\3', self.version)

    @property
    def build(self) -> int:
        """
        Extracts the build number component from the version number.
        :return: Build number integer.
        """
        return int(sub(self._version_pattern, r'\4', self.version))

    @property
    def without_build(self) -> str:
        """
        Extracts the version number without the build component.
        :return: Version number string without the build component.
        """
        return f"{self.major}.{self.minor}"
