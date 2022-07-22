from trades_log_worksheet import *
from excel_helpers import *


class OpenPositionsWS(WorkSheet):
    def __init__(self, workbook):
        WorkSheet.__init__(self, workbook, "Open Positions")
        self.trades_log = TradesLogWS(workbook)

    def update(self):
        self.update_avg_entry_price()

    def update_avg_entry_price(self):
        print("Updating average entry price...")
        trade_dict = {}
        for trade_row in self.trades_log.worksheet.iter_rows(min_row=2):
            unrealized_profit_loss = trade_row[22].value
            if unrealized_profit_loss is None:
                continue

            trades_log_stock_symbol = trade_row[2].value
            if trades_log_stock_symbol not in trade_dict.keys():
                trade_dict[trades_log_stock_symbol] = {
                    "sum_product_tot": 0,
                    "num_shares_tot": 0,
                    "cost_basis": 0,
                    "share_price": 0
                }
            trade_dict[trades_log_stock_symbol]["sum_product_tot"] += trade_row[4].value * trade_row[5].value
            trade_dict[trades_log_stock_symbol]["num_shares_tot"] += trade_row[5].value
            trade_dict[trades_log_stock_symbol]["cost_basis"] = \
                trade_dict[trades_log_stock_symbol]["sum_product_tot"] / \
                trade_dict[trades_log_stock_symbol]["num_shares_tot"]
            trade_dict[trades_log_stock_symbol]["share_price"] = trade_row[21].value
        clear_worksheet(self.worksheet, "A2:D26")
        clear_worksheet(self.worksheet, "G2:H26")
        clear_worksheet(self.worksheet, "K2:K26")

        row_counter = 1
        for stock_symbol in trade_dict:
            row = list(self.worksheet.rows)[row_counter]
            row[0].value = stock_symbol
            row[1].value = trade_dict[stock_symbol]["cost_basis"]
            row[2].value = trade_dict[stock_symbol]["share_price"]
            row[3].value = trade_dict[stock_symbol]["num_shares_tot"]
            row_counter = row_counter + 1

        # for row in self.worksheet.iter_rows(min_row=2):
        #     if stock_symbol is None:
        #         continue
        #     print("Working on: " + stock_symbol)




