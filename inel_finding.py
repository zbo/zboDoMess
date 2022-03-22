import csv
import unittest
from os import walk

'''模式查找范围10天'''
scanrange = 10

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


if __name__ == '__main__':
    #unittest.main()
    files = get_all_files()
    data = load_data('002246.sz.csv')
    peak_index = find_peak(data)
    print(data[peak_index].high)
    print(data[peak_index].vol)
    print(len(files))
    