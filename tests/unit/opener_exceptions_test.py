# -*- coding: utf-8 -*-
import json
from os import getcwd
from os.path import join

import pytest

@pytest.fixture
def json_data():
    with open(join(getcwd(), 'tests', 'assets', 'opener_exception.json'), 'r') as file:
        return json.load(file)

def test_json_structure(json_data):
    for bug_name, info in json_data.items():
        assert info is not None

        assert 'link' in info, f"{bug_name} has no 'link' parameter"
        assert isinstance(info['link'], str)

        assert 'description' in info, f"{bug_name} has no 'description' parameter"
        assert isinstance(info['description'], str)

        assert 'mode' in info, f"{bug_name} has no 'mode' parameter"
        assert isinstance(info['mode'], list)

        assert 'directions' in info, f"{bug_name} has no 'directions' parameter"
        assert isinstance(info['directions'], list)

        assert 'files' in info, f"{bug_name} has no 'files' parameter"
        assert isinstance(info['files'], list)

def test_mode_parameter(json_data):
    for bug_name, info in json_data.items():
        assert set(info['mode']), f"Bug {bug_name} has an empty 'mode'"
        valid_modes = {"Default", "t-format"}
        assert set(info['mode']) <= valid_modes, f"{bug_name} does not have a valid 'mode'. Valid value: {valid_modes}"
