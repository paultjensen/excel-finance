import shutil
from trade_log import *

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    filename = "D:\projects\excel-finance\Jensen_Family_Finances_20220716.xlsx"
    shutil.copyfile(filename, filename + "~")
    wb = load_workbook(filename=filename)
    process_trade_log(wb)
    wb.save(filename=filename)
