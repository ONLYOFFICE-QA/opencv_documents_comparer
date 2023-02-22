# -*- coding: utf-8 -*-
import subprocess as sb

sb.call('pip install tomlkit==0.11.6', shell=True)

with open("requirements.txt", 'w') as file:
    file.write('# -*- coding: utf-8 -*-\n')
with open("poetry.lock") as t:
    import tomlkit
    lock = tomlkit.parse(t.read())
    for package in lock['package']:
        with open("requirements.txt", 'a') as file:
            file.write(f"{package['name']}=={package['version']}\n")
