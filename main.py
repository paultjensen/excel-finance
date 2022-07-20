import shutil
from excel_helpers import *
from trades_log_worksheet import *
from open_positions_worksheet import *

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    filename = "D:\projects\excel-finance\Jensen_Family_Finances_20220716.xlsx"
    shutil.copyfile(filename, filename + "~")
    wb = load_workbook(filename=filename)
    trade_log = TradesLogWS(wb)
    trade_log.update_share_prices()
    trade_log.update_max_profit_loss()
    open_positions = OpenPositionsWS(wb)
    wb.save(filename=filename)
