import os
from os.path import join

import pytest
from host_tools import File
from collections import Counter


@pytest.fixture
def file_structure():
    return File.read_json(join(os.getcwd(), 'tests', 'assets', 'quick_check_files.json'))

def test_no_duplicates(file_structure):
    all_files = []
    for ext, files in file_structure.items():
        all_files.extend(files)
    duplicates = [item for item, count in Counter(all_files).items() if count > 1]
    assert not duplicates, f"Duplicate files have been detected: {duplicates}"

def test_extensions_match_keys(file_structure):
    for ext, files in file_structure.items():
        for file in files:
            assert file.lower().endswith("." + ext), f"The file extension {file} does not match the key {ext}"

def test_no_spaces_in_keys(file_structure):
    for bug_id, bug_data in file_structure.items():
        assert ' ' not in bug_id, f"Bug ID '{bug_id}' contains spaces"
        for key in bug_data.keys():
            assert ' ' not in key, f"Key '{key}' in bug {bug_id} contains spaces"
