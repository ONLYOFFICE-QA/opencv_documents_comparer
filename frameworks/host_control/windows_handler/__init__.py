from .. import HostInfo

if HostInfo().os == 'windows':
    from .windows import Window
else:
    from .linux_window import LinuxWindow as Window
