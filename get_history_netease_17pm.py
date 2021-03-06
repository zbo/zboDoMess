# coding=utf-8
from openpyxl.styles import Border, Side
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from openpyxl import Workbook
from io import StringIO
import os
import csv
import json
import time
import comlib

import requests
from openpyxl.styles.builtins import styles

fileout = './content/history.xlsx'
url_format = 'http://quotes.money.163.com/service/chddata.html?code={0}&start=20220101&end=20221231'
# url_format = 'http://quotes.money.163.com/service/chddata.html?code={0}&start=20201115&end=20210126&fields=PCHG'

wb = Workbook()
ws = wb.active

if os.path.exists(fileout):
    os.remove(fileout)
    print('history file deleted')
else:
    print('no such file:%s' % fileout)

def gen_meta_from_xl():
    meta_all_urls = []
    meta_all_codes = []
    wb = load_workbook(filename=comlib.filein)
    high_sheet = wb[comlib.highsheet]
    index = 2
    while high_sheet['A{0}'.format(index)].value is not None:
        code = high_sheet['A{0}'.format(index)].value
        index = index + 1
        meta_all_codes.append(code)
        code_short = code.split('.')[0]
        mkt = '0'
        if code.split('.')[1].lower() == 'sz':
            mkt = '1'
        n_code = mkt + code_short
        url_formatted = url_format.format(n_code)
        meta_all_urls.append(url_formatted)
    return meta_all_urls, meta_all_codes
        


def file_exist(file_name):
    if os.path.exists('./store/{0}'.format(file_name)):
        return True
    return False


if __name__ == '__main__':
    all_urls, all_codes = gen_meta_from_xl()
    index = 0
    data_bag = []
    for u in all_urls:
        s = comlib.Stock()
        filename = all_codes[index] + '.csv'
        if not file_exist(filename):
            print(u)
            response = requests.get(u)
            f = StringIO(response.text)
            csv_reader = csv.reader(f, delimiter=',')
            f_local = open('./store/{0}'.format(filename), 'w')
            csv_writer = csv.writer(f_local)
            for line in csv_reader:
                csv_writer.writerow(line)
            f_local.close()
            time.sleep(0.5)

        f_local = open('./store/{0}'.format(filename), 'r')
        reader = csv.reader(f_local, delimiter=',')
        # print(filename)
        for row in reader:
            if reader.line_num == 1:
                continue
            s.add_Item(row[9])
            s.add_Day(row[0])
            s.name = row[2]
            s.code = all_codes[index]
        data_bag.append(s)
        f_local.close()
        index = index + 1
        print("processed {0}, current request grab {1} rows".format(
            index, len(s.each_day)))

    comlib.gen_excel(data_bag, ws)
    wb.save(fileout)
