from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side, Font
import csv,os
from openpyxl import Workbook

fileout = '/Users/zhubo/Documents/out.xlsx'
wb = load_workbook(filename=fileout)

def get_cate_from_conf():
    all_cate=[]
    with open('config/category.csv', 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            all_cate.append(row[0])
    all_cate.append('其他')
    return all_cate


def update_now(sheet):
    print(sheet.max_row)
    

if __name__ == '__main__':
    all_cate = get_cate_from_conf()
    print(all_cate)
    for cate in all_cate:
        update_now(wb[cate])
    wb.close()