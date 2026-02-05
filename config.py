# -*- coding: utf-8 -*-
from host_tools import HostInfo
from os import cpu_count
from os.path import join, expanduser

version = ''

# Path to source documents
source_docs = r'C:\db\files' if HostInfo().is_windows \
    else join(expanduser('~'), 'db', 'files') if HostInfo().is_mac \
    else join(expanduser('~'), 'db', 'files') if HostInfo().is_linux \
    else None

# Path to converted documents
converted_docs = r'C:\db\results' if HostInfo().is_windows \
    else join(expanduser('~'), 'scripts', 'opencv_documents_comparer', 'result') if HostInfo().is_mac \
    else join(expanduser('~'), 'db', 'results') if HostInfo().is_linux \
    else None

# Delay time after opening a document in the editors in seconds
delay_word = 1
delay_excel = 1
delay_power_point = 2
delay_libre = 1

# Maximum amount of time the document is waiting to be opened.
max_waiting_time = 60

# x2ttester settings
cores = cpu_count() or 4
errors_only = True
delete = True
timestamp = True
timeout = 900
# Path to MS Office
ms_office = r'C:/Program Files/Microsoft Office/root/Office16'
# Path to LibreOffice
libre_office = r'C:/Program Files/LibreOffice/program'

# To test from the names array
files_array = []
