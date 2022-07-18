from excel_helpers import *
from trade_log import *

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    wb = load_workbook(filename="D:\projects\excel-finance\Jensen_Family_Finances_20220716.xlsm")
    process_trade_log(wb)
