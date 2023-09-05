# -*- coding: utf-8 -*-
import win32con
import win32gui
import time

from rich import print
from rich.console import Console

console = Console()


class Window:
    def __init__(self):
        self.windows_handler_number = None
        self.error = False
        self.window_checker = False

    @staticmethod
    def check_window_isvisible(hwnd: int, func_name: str) -> bool:
        if win32gui.IsWindowVisible(hwnd):
            return True
        print(f"[bold red]|INFO| Windows not found when `{func_name}`. Windows number: {hwnd}")

    @staticmethod
    def set_foreground(hwnd):
        if Window.check_window_isvisible(hwnd, 'set_foreground'):
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                print(f"[red]|WARNING| Can't set foreground the window. Exception: {e}")

    @staticmethod
    def set_size(hwnd: int, left: int, top: int, right: int, bottom: int) -> None:
        if Window.check_window_isvisible(hwnd, 'set_size'):
            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
            win32gui.MoveWindow(hwnd, left, top, right, bottom, True)

    @staticmethod
    def set_size_max(hwnd: int) -> None:
        if Window.check_window_isvisible(hwnd, 'set_size_max'):
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

    @staticmethod
    def move_window(hwnd, x, y):
        if not Window.check_window_isvisible(hwnd, 'move_window'):
            return
        _, _, width, height = win32gui.GetWindowRect(hwnd)
        win32gui.SetWindowPos(hwnd, None, x, y, width, height, 0)

    @staticmethod
    def get_hwnd(class_name: str = None, window_name: str = None) -> int:
        return win32gui.FindWindow(class_name, window_name)

    @staticmethod
    def get_coordinate(hwnd: int) -> list:
        if Window.check_window_isvisible(hwnd, 'get_coordinate'):
            return [win32gui.GetWindowRect(hwnd)][0]

    def close_all_by_class_names(self, class_names: list) -> bool:
        hwnd_list = self.get_hwnd_by_class_names(class_names)
        if hwnd_list:
            for hwnd in hwnd_list:
                self.close(hwnd)
            return True
        return False

    @staticmethod
    def click_on_button(button_hwnd: int):
        try:
            win32gui.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
        except Exception as ex:
            print(ex)

    @staticmethod
    def get_window_info(hwnd: int, window_title: str, window_text: str) -> list[int]:
        def find_button(hwnd, data):
            clsname, text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
            if clsname in window_title and window_text in text:
                data.append(hwnd)

        button_handles = []
        win32gui.EnumChildWindows(hwnd, find_button, button_handles)
        return button_handles

    @staticmethod
    def get_title_text(class_names: list) -> list[str]:
        def enum_windows_callback(hwnd, info):
            if win32gui.IsWindowVisible(hwnd):
                class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
                if class_name in class_names:
                    info.append(class_name)
                    info.append(window_text)

        window_info = []
        win32gui.EnumWindows(enum_windows_callback, window_info)
        return window_info

    def handle_events(self, event_handler: object()) -> bool:
        """
        :type event_handler: object(class_names, windows_text, hwnd):
        function for additional processing of window opening events
        if the 'event_handler' function returns True,
        the function 'handle_events' will return True else False

        Example event handler function:

        def event_handler(class_name, windows_text, hwnd):
            match [class_name, windows_text]:
                case ["<class name>", "<window title text>"]:
                    perform actions with the window
                    return True
        """

        def enum_windows_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                class_name, window_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
                if event_handler(class_name, window_text, hwnd) is True:
                    self.error = True

        self.error = False
        win32gui.EnumWindows(enum_windows_callback, None)
        return self.error

    @staticmethod
    def get_hwnd_by_class_names(class_names: list) -> list[int]:
        hwnd_list = []

        def enum_windows_callback(hwnd: int, hwnds: list):
            if win32gui.IsWindowVisible(hwnd):
                if win32gui.GetClassName(hwnd) in class_names:
                    hwnds.append(hwnd)

        win32gui.EnumWindows(enum_windows_callback, hwnd_list)
        return hwnd_list

    def wait_until_open(
            self,
            titles: list,
            window_title_text: str = None,
            event_handler: object() = None,
            max_waiting_time: int = 60,
            max_time_run_windows: int = 10
    ) -> int | str:
        """
        :param window_title_text:
        :param max_time_run_windows: maximum waiting time in sec
        :param max_waiting_time: maximum number of attempts to wait for the window to open
        :param titles: window titles
        :type event_handler: object(class_names, windows_text, hwnd):
        function for additional processing of window opening events (optional)
        if the 'event_handler' function returns True,
        the waiting for the opening will be interrupted and the function will return 'ERROR'.

        Example event handler function:

        def event_handler(class_name, windows_text, hwnd):
        match [class_name, windows_text]:
            case ["<class name>", "<window title text>"]:
                perform actions with the window
                return True
        """

        self.windows_handler_number = None
        self.error = False
        start_time = time.perf_counter()

        def enum_windows_callback(hwnd: int, windows_title: list):
            if win32gui.IsWindowVisible(hwnd):
                class_names, windows_text = win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)
                if class_names in windows_title:
                    self.window_checker = True
                    if window_title_text and window_title_text in windows_text:
                        self.windows_handler_number = hwnd
                    elif not window_title_text and windows_text != '':
                        self.windows_handler_number = hwnd
                    if event_handler and event_handler(class_names, windows_text, hwnd=hwnd):
                        self.error = True
                        return

        wait_time = 0
        with console.status("[red]Waiting for opening") as status:
            while wait_time < max_waiting_time:
                wait_time = time.perf_counter() - start_time
                status.update(f"[red] Waiting for opening: {wait_time:.02f}/{max_waiting_time} second. ")
                self.window_checker = False
                win32gui.EnumWindows(enum_windows_callback, titles)
                if self.error:
                    return 'ERROR'
                elif isinstance(self.windows_handler_number, int):
                    return self.windows_handler_number
                elif not self.window_checker:
                    if wait_time > max_time_run_windows:
                        console.print(f"[bold red]|ERROR| Window not opened")
                        return 'NOT_OPENED'
        console.print(f"[bold red]|ERROR| Too long to open file")
        return 'TOO_LOONG'

    @staticmethod
    def close(hwnd: int) -> None:
        if Window.check_window_isvisible(hwnd, 'close'):
            try:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception as e:
                print(f"[red]|WARNING| Can't close the window. Exception: {e}")
