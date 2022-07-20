from worksheet import WorkSheet
import yfinance as yf


class OpenPositionsWS(WorkSheet):
    def __init__(self, workbook):
        WorkSheet.__init__(self, workbook, "Trades Log")
