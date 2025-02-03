# -*- coding: utf-8 -*-
from os import getcwd
from os.path import join

import pytest
from host_tools import File
from collections import Counter


@pytest.fixture
def json_data():
    return File.read_json(join(getcwd(), 'tests', 'assets', 'conversion_exception.json'))

def test_json_structure(json_data):
    for bug_name, info in json_data.items():
        assert info is not None

        assert 'link' in info, f"{bug_name} has no 'link' parameter"
        assert isinstance(info['link'], str)

        assert 'description' in info, f"{bug_name} has no 'description' parameter"
        assert isinstance(info['description'], str)

        assert 'os' in info, f"{bug_name} has no 'mode' parameter"
        assert isinstance(info['os'], list)

        assert 'directions' in info, f"{bug_name} has no 'directions' parameter"
        assert isinstance(info['directions'], list)

        assert 'files' in info, f"{bug_name} has no 'files' parameter"
        assert isinstance(info['files'], list)

def test_os_parameter(json_data):
    for bug_name, info in json_data.items():
        valid_modes = {"linux", "windows", "mac"}
        assert set(info['os']) <= valid_modes, f"{bug_name} does not have a valid 'os'. Valid os: {valid_modes}"

def test_directions_parameter(json_data):
    for bug_id, bug_data in json_data.items():
        for direction in bug_data['directions']:
            assert direction.islower(), f"Direction '{direction}' in bug {bug_id} is not lowercase"
            assert '-' in direction, f"Direction '{direction}' in bug {bug_id} does not contain '-'"
            assert '.' not in direction, f"Direction '{direction}' in bug {bug_id} contains '.'"

def test_files_duplicates(json_data):
    for bug_id, bug_data in json_data.items():
        assert set(bug_data['files']), f"The 'files' in bug {bug_id} cannot be empty"
        duplicates = [item for item, count in Counter(bug_data['files']).items() if count > 1]
        assert not duplicates, f"Duplicate files have been detected: {duplicates} in bug: {bug_id}"

def test_no_spaces_in_keys(json_data):
    for bug_id, bug_data in json_data.items():
        assert ' ' not in bug_id, f"Bug ID '{bug_id}' contains spaces"
        for key in bug_data.keys():
            assert ' ' not in key, f"Key '{key}' in bug {bug_id} contains spaces"
