# -*- coding: utf-8 -*-
import config

from os.path import basename, splitext
from time import sleep
from rich import print
from frameworks.decorators import retry
from frameworks.editors.editor import Editor
from host_tools import File, HostInfo, Process


if HostInfo().os == "windows":
    from host_tools import Window


class Document:
    """
    Class representing a document and its interactions with an editor.

    Attributes:
        editor (Editor): The editor object used for working with the document.
        formats: Supported formats by the editor.
        events: Event handler for the editor.
        editor_class_names: Class names associated with the editor window.
        delay_after_open (float): Delay time after opening the document.
        max_waiting_time: Maximum waiting time for operations.
        window: Window object for interacting with the operating system.
    """

    def __init__(self, editor: Editor):
        """
        Initializes the Document object.
        :param editor: Editor object used for working with the document.
        """
        self.editor = editor
        self.formats = editor.formats
        self.events = editor.events_handler()
        self.editor_class_names = editor.events_handler().window_class_names
        self.delay_after_open: int | float = editor.delay_after_open
        self.max_waiting_time = config.max_waiting_time
        self.window = Window() if HostInfo().os == "windows" else None

    def open(self, file_path: str) -> int | str:
        """
        Opens a document using the editor.
        :param file_path: Path to the document file.
        :return: Identifier of the opened document or error message.
        """
        self.editor.open(file_path)
        return self._wait_until_open(file_path)

    @retry(max_attempts=10, interval=0.2, silence=True, exception=False)
    def _wait_until_open(self, file_path) -> int | str:
        """
        Waits until the document is opened.
        :param file_path: Path to the document file.
        :return: Identifier of the opened document or error message.
        """
        return self.window.wait_until_open(
            titles=self.editor_class_names,
            window_title_text=f"{basename(file_path).replace(splitext(file_path)[1], '')}",
            max_waiting_time=self.max_waiting_time,
            event_handler=self.events.when_opening,
            max_time_run_windows=10
        )

    @retry(max_attempts=10, interval=1, silence=True, exception=False)
    def check_errors(self) -> bool:
        """
        Checks for errors while opening the document.
        :return: True if there are errors, False otherwise.
        """
        return self.window.handle_events(self.events.when_opening)

    def close_all_window(self) -> bool:
        """
        Closes all windows associated with the document.
        :return: True if windows are closed successfully, False otherwise.
        """
        return self.window.close_all_by_class_names(self.editor_class_names)

    def close(self, hwnd: int = None) -> bool:
        """
        Closes the document.
        :param hwnd: Identifier of the document window.
        :return: True if the document is closed successfully, False otherwise.
        """
        self.editor.close(hwnd) if hwnd else ...
        for i in range(15):
            if not self.close_all_window():
                return True
            print(f"[green] |INFO| Try to close windows: {i + 1}"), sleep(0.2)
        Process.terminate(self.editor.process_names)
        return False

    def make_screenshots(self, hwnd: int, screen_path: str, page_amount: int | dict) -> None:
        """
        Makes screenshots of the document.
        :param hwnd: Identifier of the document window.
        :param screen_path: Path to save the screenshots.
        :param page_amount: Amount of pages to screenshot.
        :return: None
        """
        self.editor.set_size(hwnd)
        return self.editor.make_screenshots(hwnd, screen_path, page_amount)

    def page_amount(self, file_path: str) -> int | dict:
        """
        Retrieves the amount of pages in the document.
        :param file_path: Path to the document file.
        :return: Amount of pages or dictionary with page information.
        """
        return self.editor.page_amount(file_path)

    def delete(self, file_path: str, ) -> bool:
        """
        Deletes the document file.
        :param file_path: Path to the document file.
        :return: True if the document is deleted successfully, False otherwise.
        """
        try:
            self._deleter(file_path)
            return True
        except Exception as e:
            print(f"[bold red]|ERROR| Can't delete document. Exception: {e}")
            Process.terminate(self.editor.process_names)
            return False

    @staticmethod
    @retry(max_attempts=10, interval=1, silence=True)
    def _deleter(file_path: str) -> None:
        """
        Deletes the document file.
        :param file_path: Path to the document file.
        :return: None
        """
        File.delete(file_path, stdout=False)
