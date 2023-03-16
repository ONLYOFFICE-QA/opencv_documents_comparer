# -*- coding: utf-8 -*-
from functools import wraps
from time import perf_counter

from rich import print


def singleton(class_):
    __instances = {}

    @wraps(class_)
    def getinstance(*args, **kwargs):
        if class_ not in __instances:
            __instances[class_] = class_(*args, **kwargs)
        return __instances[class_]

    return getinstance


def highlighter(color='green'):
    def line_printer_inner(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            print(f'[{color}]-' * 90)
            result = func(*args, **kwargs)
            print(f'[{color}]-' * 90)
            return result

        return wrapper

    return line_printer_inner


def timer_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = perf_counter()
        result = func(*args, **kwargs)
        print(f"Time existing func {func.__name__}: {(perf_counter() - start_time):.02f}")
        return result

    return wrapper


def memoize(func):
    cache = {}

    @wraps(func)
    def wrapper(*args, **kwargs):
        key = (args, tuple(kwargs.items()))
        if key not in cache:
            cache[key] = func(*args, **kwargs)
        return cache[key]

    return wrapper
