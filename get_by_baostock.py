import baostock as bs
import pandas as pd
import csv,os
import comlib
from openpyxl import Workbook

fileout = '/Users/zhubo/Documents/history_bao.xlsx'
wb = Workbook()
ws = wb.active

if os.path.exists(fileout):
    os.remove(fileout)
    print('bao history file deleted')
else:
    print('no such file:%s'%fileout)

def get_single(code):
    rs = bs.query_history_k_data_plus(code,
                                      "date,code,open,high,low,close,preclose,volume,amount,adjustflag,turn,tradestatus,pctChg,isST",
                                      start_date='2022-03-01', end_date='2022-12-31',
                                      frequency="d", adjustflag="3")
    print('query_history_k_data_plus respond error_code:' + rs.error_code)
    print('query_history_k_data_plus respond  error_msg:' + rs.error_msg)
    data_list = []
    while (rs.error_code == '0') & rs.next():
        data_list.append(rs.get_row_data())
    result = pd.DataFrame(data_list, columns=rs.fields)
    return result


def get_codes():
    r = []
    with open('reduced.csv', 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            code = row[0].split('.')
            r.append('{0}.{1}'.format(code[1], code[0]))
    return r


if __name__ == '__main__':
    lg = bs.login()
    print('login respond error_code:' + lg.error_code)
    print('login respond  error_msg:' + lg.error_msg)
    codes = get_codes()
    data_bag = []
    for code in codes:
        result = get_single(code)
        # result.to_csv("D:\\history_A_stock_k_data.csv", index=False)
        s = comlib.fill_stock(result)
        data_bag.append(s)
        # print(result)
    bs.logout()
    comlib.gen_excel(data_bag, ws)
    wb.save(fileout)
