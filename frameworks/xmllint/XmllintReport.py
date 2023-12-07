# -*- coding: utf-8 -*-
from datetime import datetime
from os.path import join

from frameworks.StaticData import StaticData
from frameworks.decorators import singleton
from host_tools import HostInfo, Dir
from frameworks.editors.onlyoffice import VersionHandler, X2t
from frameworks.report import Report


@singleton
class XmllintReport(Report):
    def __init__(self):
        super().__init__()
        self.report_dir = StaticData.reports_dir()
        self.time_pattern = f"{datetime.now().strftime('%H_%M_%S')}"
        self.x2t_dir = StaticData.core_dir()

    def report_path(self):
        version = X2t.version(self.x2t_dir)
        report_dir = join(self.report_dir, VersionHandler(version).without_build, HostInfo().os)
        Dir.create(report_dir)
        return join(report_dir, f"{version}_{self.time_pattern}.csv")
