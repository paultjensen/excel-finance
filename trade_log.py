import math
import yfinance as yf
from excel_helpers import *


def process_trade_log(workbook):
    ws = get_worksheet(workbook, 'Trades Log')
    for row in ws.iter_rows(min_row=2):
        stock_symbol = row[2].value
        entry_price = row[4].value
        start_date = row[3].value
        end_date = row[14].value
        max_drawdown = row[17].value
        max_profit = row[18].value
        if stock_symbol is None:
            continue
        if max_drawdown is not None and max_profit is not None:
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

