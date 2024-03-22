# -*- coding: utf-8 -*-
from functools import wraps
from multiprocessing import Process
from time import perf_counter, sleep

from rich import print


def singleton(class_):
    """
    Decorator to ensure a class has only one instance.
    :param class_: Class to decorate.
    :return: Instance of the class.
    """
    __instances = {}

    @wraps(class_)
    def getinstance(*args, **kwargs):
        if class_ not in __instances:
            __instances[class_] = class_(*args, **kwargs)
        return __instances[class_]

    return getinstance


def retry(max_attempts: int = 3, interval: int | float = 0, silence: bool = False, exception: bool = True):
    """
    Decorator to retry a function multiple times.
    :param max_attempts: Maximum number of attempts.
    :param interval: Interval between attempts.
    :param silence: If True, suppresses error messages.
    :param exception: If True, raises exception after all attempts fail.
    :return: Result of the decorated function.
    """
    def wrapper(func):
        @wraps(func)
        def inner(*args, **kwargs):
            for i in range(max_attempts):
                try:
                    result = func(*args, **kwargs)
                except Exception as e:
                    print(f"[cyan] |INFO| Exception when '{func.__name__}'. Try: {i + 1} of {max_attempts}.")
                    print(f"[red]|WARNING| Error: {e}[/]") if not silence else ...
                    sleep(interval)
                else:
                    return result
            print(f"[bold red]|ERROR| The function: '{func.__name__}' failed in {max_attempts} attempts.")
            if exception:
                raise

        return inner

    return wrapper


def highlighter(color: str = 'green'):
    """
    Decorator to print highlighting before and after the function execution.
    :param color: Color for highlighting.
    :return: Result of the decorated function.
    """
    def wrapper(func):
        @wraps(func)
        def inner(*args, **kwargs):
            print(f'[{color}]-' * 90)
            result = func(*args, **kwargs)
            print(f'[{color}]-' * 90)
            return result

        return inner

    return wrapper


def async_processing(target):
    """
    Decorator to run a function asynchronously using multiprocessing.
    :param target: Target function to run asynchronously.
    :return: Result of the decorated function.
    """
    def wrapper(func):
        @wraps(func)
        def inner(*args, **kwargs):
            process = Process(target=target)
            process.start()
            result = func(*args, **kwargs)
            process.terminate()
            return result

        return inner

    return wrapper


def timer(func):
    """
    Decorator to measure the execution time of a function.
    :param func: Function to measure execution time.
    :return: Result of the decorated function.
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = perf_counter()
        result = func(*args, **kwargs)
        print(f"[green]|INFO| Time existing the function `{func.__name__}`: {(perf_counter() - start_time):.02f}")
        return result

    return wrapper


def memoize(func):
    """
    Decorator to cache the result of a function.
    :param func: Function to memoize.
    :return: Result of the decorated function.
    """
    __cache = {}

    @wraps(func)
    def wrapper(*args, **kwargs):
        key = (args, tuple(kwargs.items()))
        if key not in __cache:
            __cache[key] = func(*args, **kwargs)
        return __cache[key]

    return wrapper
