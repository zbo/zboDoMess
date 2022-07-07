# coding='utf-8'

import csv
import unittest
from os import walk
import comlib


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
    for (dirpath, dirnames, filenames) in walk('./bao/'):
        f.extend(filenames)
        break
    return f


def load_data(filename):
    f_local = open('./bao/{0}'.format(filename), 'r')
    reader = csv.reader(f_local, delimiter=',')
    his_stock = []
    for row in reader:
        s = Stock()
        if reader.line_num == 1:
            continue
        # bao stock return no names
        s.name = row[1]
        s.code = row[1]
        s.date = row[0]
        s.change = float(row[5])-float(row[6])
        s.vol = row[7]
        s.high = row[3]
        s.low = row[4]
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

    def test_bllh(self):
        data = load_data('600322.sh.csv')
        peak_index = find_peak(data)
        print(peak_index)

    @unittest.skip
    def test_upper(self):
        data = load_data('sh.600007.csv')
        peak_index = find_peak(data)
        self.assertEqual(peak_index, 0)

    @unittest.skip
    def test_isupper(self):
        data = load_data('002246.sz.csv')
        peak_index = find_peak(data)
        self.assertEqual(peak_index, 3)


def range_meet(data_range, high_vol):
    meet = True
    for d in data_range:
        if d.vol == '':
            continue
        if float(d.vol) > float(high_vol) and float(d.change) < 0:
            meet = False
            break
    return meet

def logic_fx():
    code_list = []
    files = get_all_files()
    #print('search fx total {0} files found in bao'.format(len(files)))
    for f in files:
        data = load_data(f)
        data.reverse()
        hit_bottom = data[1].low <= data[2].low and data[1].high <= data[2].high
        protect_bottom = data[0].low >= data[1].low
        if hit_bottom and protect_bottom:
            arr = f.split('.')
            code_list.append('{0}.{1}'.format(arr[0], arr[1]))
    #print('total {0} item meet fx condition'.format(len(code_list)))
    return code_list

def logic():
    code_list = []
    files = get_all_files()
    print('search sl total {0} files found in bao'.format(len(files)))
    for f in files:
        data = load_data(f)
        data.reverse()
        peak_index = find_peak(data)
        if remove_meet(data, peak_index):
            continue
        data_range = data[:peak_index]
        high_vol = data[peak_index].vol
        if range_meet(data_range, high_vol):
            arr = f.split('.')
            code_list.append('{0}.{1}'.format(arr[0], arr[1]))
            # print(data[0].name)
    reverse_dot = []
    for item in code_list:
        arr = item.split('.')
        reverse_dot.append('{0}.{1}'.format(arr[1],arr[0]))
    return reverse_dot


def check_no_zt(data, ran):
    for i in range(ran):
        if data[i].change > 0.9:
            return False
    return True


def remove_meet(data, peak_index):
    too_long = peak_index > comlib.sl_scan_range or peak_index == 0
    peak_dark = float(data[peak_index].change) < 0
    nh = peak_index == 0
    no_zt_ten_days = check_no_zt(data,10)
    return too_long or peak_dark or no_zt_ten_days or nh


if __name__ == '__main__':
    # unittest.main()
    all_code = logic_fx()
    for code in all_code:
        print(code)
        # print(code.split('.')[0])
