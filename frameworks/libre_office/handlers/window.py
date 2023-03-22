import win32con
import win32gui


class Window:
    def __init__(self, windows_handler_number):
        self.windows_handler_number = windows_handler_number

    def set_size(self):
        win32gui.ShowWindow(self.windows_handler_number, win32con.SW_NORMAL)
        win32gui.MoveWindow(self.windows_handler_number, 0, 0, 2200, 1420, True)
        win32gui.SetForegroundWindow(self.windows_handler_number)

    def get_coordinate(self):
        return [win32gui.GetWindowRect(self.windows_handler_number)][0]
