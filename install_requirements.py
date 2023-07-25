# -*- coding: utf-8 -*-
import os
import re
import tomllib
from os.path import join
from platform import system
from subprocess import call

requirements_file = join(os.getcwd(), 'requirements.txt')
poetry_toml = join(os.getcwd(), 'pyproject.toml')
exceptions = ['python']


def read_package_dict(package_dict: dict) -> tuple[bool, str]:
    markers = package_dict.get('markers')
    version = package_dict.get('version')
    if markers and markers.split()[0] == 'sys_platform':
        if markers.split()[-1] == _os():
            if version:
                return True, re.sub(r'[*^]', '', version)
            return True, ''
    return False, ''

def create_requirements() -> None:
    _write('# -*- coding: utf-8 -*-\n')
    with open(poetry_toml, 'rb') as f:
        for package, version in  tomllib.load(f)['tool']['poetry']['dependencies'].items():
            if package.lower() not in exceptions:
                if isinstance(version, dict):
                    package_on, version = read_package_dict(version)
                    if package_on is False:
                        continue
                else:
                    version = re.sub(r'[*^]', '', version)
                _write(f"{package}{('==' + version) if version else ''}\n", 'a')

def install_requirements():
    call(f'pip install -r {requirements_file}', shell=True)

def _os() -> str:
    match system().lower():
        case 'linux':
            return 'linux'
        case 'darwin':
            return 'darwin'
        case 'windows':
            return 'win32'
        case definition_os:
            print(f"[bold red]|WARNING| Error defining os: {definition_os}")

def _write(text: str, mode: str = 'w') -> None:
    with open(requirements_file, mode=mode) as file:
        file.write(text)



if __name__ == "__main__":
    create_requirements()
    install_requirements()
