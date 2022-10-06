import baostock as bs
import pandas as pd
import csv, os
import comlib
from openpyxl import Workbook

fileout = './content/history_bao.xlsx'
wb = Workbook()
ws = wb.active

if os.path.exists(fileout):
    os.remove(fileout)
    print('bao history file deleted')
else:
    print('no such file:%s' % fileout)


def get_single(code):
    rs = bs.query_history_k_data_plus(code,
                                      "date,code,open,high,low,close,preclose,volume,amount,adjustflag,turn,tradestatus,pctChg,isST",
                                      start_date='2022-09-01', end_date='2022-12-31',
                                      frequency="d", adjustflag="3")
    data_list = []
    while (rs.error_code == '0') & rs.next():
        data_list.append(rs.get_row_data())
    result = pd.DataFrame(data_list, columns=rs.fields)
    return result


if __name__ == '__main__':
    lg = bs.login()
    print('login respond error_code:' + lg.error_code)
    print('login respond  error_msg:' + lg.error_msg)
    codes = comlib.get_orign_codes_from_xl()
    sd_codes = comlib.get_sd_codes_from_xl()
    codes = codes + sd_codes
    data_bag = []
    index = 1
    for code in codes:
        filepath = './bao/{0}.csv'.format(code)
        if os.path.exists(filepath):
            result = df = pd.read_csv(filepath)
        else:
            result = get_single(code)
            result.to_csv(filepath, index=False)
        s = comlib.fill_stock(result)
        rs = bs.query_stock_basic(code=code)
        if len(rs.data) == 0:
            s.name = 'query_failed'
        else:
            s.name = rs.data[0][1]
        data_bag.append(s)
        print('process {0} for {1}'.format(index, code))
        index = index + 1
    bs.logout()
    comlib.gen_excel(data_bag, ws)
    wb.save(fileout)
