# -*- coding: utf-8 -*-
import time

from Xlib import display, X
from Xlib.protocol import event
from rich.console import Console

console = Console()


class LinuxWindow:
    def __init__(self):
        self.window_id = None
        self.display_obj = display.Display()
        self.root = self.display_obj.screen().root

    def set_foreground(self, hwnd):
        active_window_event = event.ClientMessage(
            window=hwnd,
            client_type=self.display_obj.intern_atom('_NET_ACTIVE_WINDOW'),
            data=(32, [X.CurrentTime, 0, 0, 0, 0]),
        )
        self.root.send_event(active_window_event, event_mask=X.SubstructureRedirectMask)

    def get_hwnd(self, class_name):
        for window in self.traverse_windows(self.root):
            window_class = window.get_wm_class()
            if window_class and class_name == window_class:
                return window.id

    def traverse_windows(self, window):
        windows = []
        children = window.query_tree().children
        for child in children:
            windows.extend(self.traverse_windows(child))
        windows.append(window)
        return windows

    def wait_until_open(self, class_name, timeout: int = 60):
        start_time = time.time()
        with console.status('') as status:
            while time.time() - start_time < timeout:
                windows_id = self.get_hwnd(class_name)
                if windows_id:
                    return windows_id
                status.update(f"Wait until open window: {(time.time() - start_time):.02f}/{timeout}")
        return False
