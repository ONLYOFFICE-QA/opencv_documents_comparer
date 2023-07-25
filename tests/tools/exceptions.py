# -*- coding: utf-8 -*-
import itertools
import json


class Exceptions:
    @staticmethod
    def file_names(exceptions: json) -> list:
        return list(itertools.chain(*[i[1]['files'] for i in exceptions.items()]))

    @staticmethod
    def find_title(name: str, exceptions: json) -> str:
        for i in exceptions:
            if name in i[1]:
                return i[0]
