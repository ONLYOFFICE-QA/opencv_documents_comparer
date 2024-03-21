# -*- coding: utf-8 -*-
import platform

from .onlyoffice import X2tTester, X2tTesterConfig
from .onlyoffice import Core

if platform.system().lower() == 'windows':
    from .libre_office import LibreOffice
    from .microsoft_office import Excel, PowerPoint, Word
    from .document import Document
