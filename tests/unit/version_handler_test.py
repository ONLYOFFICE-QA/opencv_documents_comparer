# -*- coding: utf-8 -*-
import os
import sys

import pytest
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))
from frameworks.editors.onlyoffice.handlers.VersionHandler import VersionHandler


def test_version_initialization():
    handler = VersionHandler("1.2.3.4")
    assert handler.version == "1.2.3.4"

def test_version_setter_valid():
    handler = VersionHandler("0.0.0.0")
    handler.version = "1.2.3.4"
    assert handler.version == "1.2.3.4"

def test_version_setter_invalid():
    handler = VersionHandler("1.2.3.4")
    with pytest.raises(ValueError):
        handler.version = "1.2.3"

def test_major():
    handler = VersionHandler("12.34.56.78")
    assert handler.major == "12.34"

def test_minor():
    handler = VersionHandler("12.34.56.78")
    assert handler.minor == 56

def test_build():
    handler = VersionHandler("12.34.56.78")
    assert handler.build == 78

def test_without_build():
    handler = VersionHandler("12.34.56.78")
    assert handler.without_build == "12.34.56"

def test_version_setter_invalid_format():
    with pytest.raises(ValueError):
        VersionHandler("12.34.56")

def test_version_setter_invalid_characters():
    with pytest.raises(ValueError):
        VersionHandler("12.a.56.78")
