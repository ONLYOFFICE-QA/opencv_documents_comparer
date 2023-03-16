# -*- coding: utf-8 -*-
import tomllib

with open('poetry.lock', 'rb') as f:
    lock = tomllib.load(f)

with open("requirements.txt", 'w') as file:
    file.write('# -*- coding: utf-8 -*-\n')
for package in lock['package']:
    with open("requirements.txt", 'a') as file:
        file.write(f"{package['name']}=={package['version']}\n")
