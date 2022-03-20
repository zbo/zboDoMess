from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side
import csv,os
from openpyxl import Workbook

filein = '/Users/zhubo/Documents/中美对话周期.xlsx'
fileout = '/Users/zhubo/Documents/out.xlsx'
wb_out = Workbook()
wb = load_workbook(filename=filein)
sheet_ranges = wb['虎年高度']
category = []
existing = []
others = []

if os.path.exists(fileout):
    os.remove(fileout)
    print('file deleted')
else:
    print('no such file:%s'%fileout)


def get_categoty():
    global category, cate, index
    category = set()
    while sheet_ranges['c{0}'.format(index)].value is not None:
        cate = sheet_ranges['c{0}'.format(index)].value
        category.add(cate)
        index = index + 1
    return category


def get_others():
    global existing
    with open('config/category.csv', 'r') as f:
        existing = []
        all_array = []
        reader = csv.reader(f)
        for row in reader:
            existing.append(row)
            all_array.extend(row[1:])
    for s in category:
        if s in all_array:
            continue
        others.append(s)
    return others


def get_cate(cate):
    for item in existing:
        if cate in item:
            return item[0]
    return '其他'


def copy_row(target_sheet, source_sheet, index):
    target_row = target_sheet.max_row+1
    for c in range(1, source_sheet.max_column+1):
        target_cell = target_sheet.cell(row=target_row, column=c)
        source_cell = source_sheet.cell(row=index, column=c)
        target_cell.value = source_cell.value
        target_cell.fill = PatternFill("solid", fgColor=source_cell.fill.fgColor)
        thin = Side(border_style="thin", color="000000")
        target_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        if target_cell.column_letter not in ['A','B','C']:
            target_sheet.column_dimensions[target_cell.column_letter].width = 3


def copy_title(target_sheet, source_sheet):
    for c in range(1, source_sheet.max_column + 1):
        target_cell = target_sheet.cell(row=1, column=c)
        source_cell = source_sheet.cell(row=1, column=c)
        target_cell.value = source_cell.value
        target_cell.fill = PatternFill("solid", fgColor=source_cell.fill.fgColor)
        thin = Side(border_style="thin", color="000000")
        target_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)


if __name__ == '__main__':
    index = 2
    category = get_categoty()
    # print(category)
    others = get_others()
    # print(others)
    for item in existing:
        if '其他' not in wb_out.sheetnames:
            ws = wb_out.create_sheet('其他')
        if item[0] not in wb_out.sheetnames:
            ws = wb_out.create_sheet(title=item[0])
    wb_out.save(fileout)

    index = 2
    while sheet_ranges['c{0}'.format(index)].value is not None:
        cate = sheet_ranges['c{0}'.format(index)].value
        fix_cate = get_cate(cate)
        copy_row(wb_out[fix_cate], sheet_ranges, index)
        index = index + 1
    for name in wb_out.sheetnames:
        copy_title(wb_out[name],sheet_ranges)
    sheet1 = wb_out['Sheet']
    wb_out.remove(sheet1)
    wb_out.save(fileout)