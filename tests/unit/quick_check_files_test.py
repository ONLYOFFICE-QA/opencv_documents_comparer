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
