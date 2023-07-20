# -*- coding: utf-8 -*-
from os.path import basename, splitext
from time import sleep

from rich import print

import config
from frameworks.decorators import retry
from frameworks.editors.editor import Editor
from frameworks.host_control import FileUtils
from frameworks.host_control import HostInfo

if HostInfo().os == "windows":
    from frameworks.host_control import Window


class Document:

    def __init__(self, editor: Editor):
        self.editor = editor
        self.formats = editor.formats
        self.events = editor.events_handler()
        self.editor_class_names = editor.events_handler().window_class_names
        self.delay_after_open: int | float = editor.delay_after_open
        self.max_waiting_time = config.max_waiting_time
        self.window = Window()

    def open(self, file_path: str) -> int | str:
        self.editor.open(file_path)
        return self._wait_until_open(file_path)

    @retry(max_attempts=10, interval=0.2, silence=True, exception=False)
    def _wait_until_open(self, file_path) -> int | str:
        return self.window.wait_until_open(
            titles=self.editor_class_names,
            window_title_text=basename(file_path).replace(splitext(file_path)[1], ''),
            max_waiting_time=self.max_waiting_time,
            event_handler=self.events.when_opening,
            max_time_run_windows=10
        )

    @retry(max_attempts=10, interval=1, silence=True, exception=False)
    def check_errors(self) -> bool:
        return self.window.handle_events(self.events.when_opening)

    def close_all_window(self) -> bool:
        return self.window.close_all_by_class_names(self.editor_class_names)

    def close(self, hwnd: int = None) -> bool:
        self.editor.close(hwnd) if hwnd else ...
        for i in range(15):
            if not self.close_all_window():
                return True
            print(f"[green] |INFO| Try to close windows: {i + 1}"), sleep(0.2)
        FileUtils.terminate_process(self.editor.process_names)
        return False

    def make_screenshots(self, hwnd: int, screen_path: str, page_amount: int | dict) -> None:
        self.editor.set_size(hwnd)
        return self.editor.make_screenshots(hwnd, screen_path, page_amount)

    def page_amount(self, file_path: str) -> int | dict:
        return self.editor.page_amount(file_path)

    def delete(self, file_path: str, ) -> bool:
        try:
            self._deleter(file_path)
            return True
        except Exception as e:
            print(f"[bold red]|ERROR| Can't delete document. Exception: {e}")
            FileUtils.terminate_process(self.editor.process_names)
            return False

    @staticmethod
    @retry(max_attempts=10, interval=1, silence=True)
    def _deleter(file_path: str) -> None:
        FileUtils.delete(file_path, silence=True)
