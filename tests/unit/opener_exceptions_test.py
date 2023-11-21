# -*- coding: utf-8 -*-
import json
from os import getcwd
from os.path import join

import pytest

@pytest.fixture
def json_data():
    with open(join(getcwd(), 'tests', 'assets', 'opener_exception.json'), 'r') as file:
        return json.load(file)

def test_info_empty(json_data):
    for _, info in json_data.items():
        assert info is not None
        assert 'link' in info
        assert 'description' in info
        assert 'mode' in info
        assert isinstance(info['mode'], list)
        assert 'directions' in info
        assert isinstance(info['directions'], list)
        assert 'files' in info
        assert isinstance(info['files'], list)

def test_info_mode(json_data):
    for _, info in json_data.items():
        valid_modes = {"Default", "t-format"}
        assert set(info['mode']) <= valid_modes
