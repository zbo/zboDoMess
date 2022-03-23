#coding='utf-8'

import csv
import unittest
from os import walk

'''模式查找范围10天'''
scanrange = 10
image_template = '<figure class="third"><img src="http://image.sinajs.cn/newchart/daily/n/{0}.gif" width="50%">' \
                 '<img src="http://image.sinajs.cn/newchart/min/n/{1}.gif" width="50%"></figure>'
link_template = 'http://vip.stock.finance.sina.com.cn/corp/go.php/vCB_AllNewsStock/symbol/{0}.phtml'
link_temp2 = 'http://vip.stock.finance.sina.com.cn/corp/go.php/vCI_CorpOtherInfo/stockid/{0}/menu_num/5.phtml'
link_taogu = 'https://www.taoguba.com.cn/quotes/{0}'


class Stock:
    def __init__(self):
        self.name = ''
        self.code = ''
        self.change = ''
        self.close = ''
        self.open = ''
        self.vol = ''
        self.date = ''
        self.high = ''


def get_all_files():
    f = []
    for (dirpath, dirnames, filenames) in walk('./store/'):
        f.extend(filenames)
        break
    return f


def load_data(filename):
    f_local = open('./store/{0}'.format(filename), 'r')
    reader = csv.reader(f_local, delimiter=',')
    his_stock = []
    for row in reader:
        s = Stock()
        if reader.line_num == 1:
            continue
        s.name = row[2]
        s.code = row[1]
        s.date = row[0]
        s.change = row[9]
        s.vol = row[11]
        s.high = row[4]
        his_stock.append(s)
    return his_stock


def find_peak(data):
    peak_index = 0
    high_value = 0
    for index, d in enumerate(data):
        if float(d.high) > high_value:
            high_value = float(d.high)
            peak_index = index
    return peak_index


class TestStringMethods(unittest.TestCase):

    def test_upper(self):
        data = load_data('000524.sz.csv')
        peak_index = find_peak(data)
        self.assertEqual(peak_index, 0)

    def test_isupper(self):
        data = load_data('002246.sz.csv')
        peak_index = find_peak(data)
        self.assertEqual(peak_index, 3)


def range_meet(data_range, high_vol):
    meet = True
    for d in data_range:
        if float(d.vol) > float(high_vol) and float(d.change) < 0:
            meet = False
            break
    return meet


def logic():
    code_list=[]
    files = get_all_files()
    print(len(files))
    for f in files:
        data = load_data(f)
        peak_index = find_peak(data)
        if peak_index > 3 or peak_index == 0:
            continue
        data_range = data[:peak_index]
        high_vol = data[peak_index].vol
        if range_meet(data_range, high_vol):
            arr = f.split('.')
            code_list.append('{0}.{1}'.format(arr[0], arr[1]))
            # print(data[0].name)
    return code_list


if __name__ == '__main__':
    # unittest.main()
    all_code = logic()
    for code in all_code:
        print(code)
        # print(code.split('.')[0])
