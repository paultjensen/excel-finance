import os
import datetime
import shutil
from open_positions_worksheet import *

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    filepath = "D:\projects\excel-finance\Jensen_Family_Finances_20220716.xlsx"
    modifiedTime = os.path.getmtime(filepath)
    timeStamp = datetime.datetime.fromtimestamp(modifiedTime).strftime("~%b%d%y%H%M%S")
    shutil.copyfile(filepath, filepath + timeStamp)
    wb = load_workbook(filename=filepath)
    trade_log = TradesLogWS(wb)
    trade_log.update()
    open_positions = OpenPositionsWS(wb)
    open_positions.update()
    wb.save(filename=filepath)
