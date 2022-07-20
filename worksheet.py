from excel_helpers import *


class WorkSheet:
    def __init__(self, workbook, wsName):
        self.workbook = workbook
        self.name = wsName
        self.worksheet = get_worksheet(self.workbook, self.name)