from excel_helpers import *
import yfinance as yf


def process_trade_log(workbook):
    ws = get_worksheet(workbook, 'Trades Log')

    for i in range(2, ws.max_row + 1):
        row = [cell.value for cell in ws[i]]
        if row[1] is not None:
            print("Working on: " + row[2] + " " + str(row[3]) + "  --->  " + str(row[14]))
            stock = yf.Ticker(row[2])
            history = stock.history(start=row[3], end=row[14])
            print(history.columns.get_level_values(0))
            for date in history.index:
                print(date)
                history_day = history[date]
                print(history_day)
