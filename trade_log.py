import math
import yfinance as yf
from excel_helpers import *


class TradeLog:
    def __init__(self, workbook):
        self.workbook = workbook
        self.wsName = "Trades Log"
        self.worksheet = get_worksheet(self.workbook, self.wsName)

    def update_share_prices(self):
        print("Updating Trade Log open trade prices...")
        for row in self.worksheet.iter_rows(min_row=2):
            stock_symbol = row[2].value
            unrealized_profit_loss = row[22].value

            if stock_symbol is None:
                continue
            if unrealized_profit_loss is None:
                continue

            print("Working on: " + stock_symbol)
            last_price = yf.Ticker(stock_symbol).info["regularMarketPrice"]
            row[21].value = last_price

    def update_max_profit_loss(self):
        print("Updating Trade Log open trade max profit/loss...")
        for row in self.worksheet.iter_rows(min_row=2):
            stock_symbol = row[2].value
            entry_price = row[4].value
            start_date = row[3].value
            end_date = row[14].value
            unrealized_profit_loss = row[22].value
            if stock_symbol is None:
                continue
            if unrealized_profit_loss is None:
                continue

            print("Working on: " + stock_symbol + " " + str(start_date) + "  --->  " + str(end_date))
            stock = yf.Ticker(stock_symbol)
            history = stock.history(start=start_date, end=end_date, back_adjust=False, auto_adjust=False)

            if history.Close.empty:
                row[17].value = 0
                row[18].value = 0
                continue

            low_prices = history.Low
            max_drawdown_perc = ((low_prices.min() - entry_price) / entry_price)
            print("\tMin Profit: %.2f \tMax Drawdown: %.2f%%" % (low_prices.min(),
                                                                 100*max_drawdown_perc))
            if not math.isnan(max_drawdown_perc):
                row[17].value = max_drawdown_perc

            high_prices = history.High
            max_profit_perc = ((high_prices.max() - entry_price) / entry_price)
            print("\tMax Price: %.2f \tMax Profit: %.2f%%" % (high_prices.max(),
                                                              100*max_profit_perc))
            if not math.isnan(max_profit_perc):
                row[18].value = max_profit_perc

