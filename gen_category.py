from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side, Font
import csv, os
from openpyxl import Workbook
import gen_finding_from_store
import gen_finding_from_bao

filein = '/Users/zhubo/Documents/大市场周期.xlsx'
fileout = '/Users/zhubo/Documents/out.xlsx'
wb_out = Workbook()
wb = load_workbook(filename=filein)
high_sheet = wb['虎年高度']
category = []
existing = []
others = []

if os.path.exists(fileout):
    os.remove(fileout)
    print('out file deleted')
else:
    print('no such file:%s' % fileout)


def get_categoty():
    index = 2
    category = set()
    while high_sheet['c{0}'.format(index)].value is not None:
        cate = high_sheet['c{0}'.format(index)].value
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
            all_array.extend(row)
    for s in category:
        if s in all_array:
            continue
        others.append(s)
    return others


def get_cate(cate):
    cates = []
    for item in existing:
        if cate in item:
            cates.append(item[0])
    if len(cates) >= 1:
        return cates
    else:
        return ['其他']


def copy_row(target_sheet, source_sheet, index):
    target_row = target_sheet.max_row + 1
    for c in range(1, source_sheet.max_column + 1):
        target_cell = target_sheet.cell(row=target_row, column=c)
        # print('index={0} column={1}'.format(index,c))
        source_cell = source_sheet.cell(row=index, column=c)
        target_cell.value = source_cell.value
        target_cell.fill = PatternFill("solid", fgColor=source_cell.fill.fgColor)
        thin = Side(border_style="thin", color="000000")
        target_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        target_cell.font = Font(name='华文楷体')
        if target_cell.column_letter not in ['A', 'B', 'C']:
            target_sheet.column_dimensions[target_cell.column_letter].width = 3


def copy_title(target_sheet, source_sheet):
    for c in range(1, source_sheet.max_column + 1):
        target_cell = target_sheet.cell(row=1, column=c)
        source_cell = source_sheet.cell(row=1, column=c)
        target_cell.value = source_cell.value
        target_cell.fill = PatternFill("solid", fgColor=source_cell.fill.fgColor)
        thin = Side(border_style="thin", color="000000")
        target_cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)


def generate_cate_sheet():
    # global index, category, others, cate, name
    index = 2
    while high_sheet['c{0}'.format(index)].value is not None:
        cate = high_sheet['c{0}'.format(index)].value
        fix_cates = get_cate(cate)
        for fxc in fix_cates:
            copy_row(wb_out[fxc], high_sheet, index)
        index = index + 1
    for name in wb_out.sheetnames:
        copy_title(wb_out[name], high_sheet)
    sheet1 = wb_out['Sheet']
    wb_out.remove(sheet1)
    wb_out.save(fileout)


def gen_sheet():
    for item in existing:
        if '其他' not in wb_out.sheetnames:
            wb_out.create_sheet('其他')
        if '高度' not in wb_out.sheetnames:
            wb_out.create_sheet('高度')
        if '缩量' not in wb_out.sheetnames:
            wb_out.create_sheet('缩量')
        if item[0] not in wb_out.sheetnames:
            wb_out.create_sheet(title=item[0])
    wb_out.save(fileout)


def meet_high_condition(source_sheet, index):
    color = []
    for c in range(1, source_sheet.max_column + 1):
        source_cell = source_sheet.cell(row=index, column=c)
        color.append(source_cell.fill.fgColor)
    max_num = 0
    total_num = 0
    for color_item in color:
        if color_item.index == 'FFFF0000':
            total_num = total_num + 1
        else:
            if total_num > max_num:
                max_num = total_num
            total_num = 0
    return max_num >= 7


def generate_top_sheet():
    for i in range(1, high_sheet.max_row):
        if i == 1:
            copy_title(wb_out['高度'], high_sheet)
        else:
            if meet_high_condition(high_sheet, i):
                copy_row(wb_out['高度'], high_sheet, i)
    wb_out.save(fileout)


def generate_sl_sheet():
    all_code_store = gen_finding_from_store.logic()
    all_code_bao = gen_finding_from_bao.logic()
    all_code = set(all_code_bao+all_code_store)
    for i in range(1, high_sheet.max_row):
        if i == 1:
            copy_title(wb_out['缩量'], high_sheet)
        else:
            if high_sheet.cell(row=i, column=1).value in all_code:
                copy_row(wb_out['缩量'], high_sheet, i)
    wb_out.save(fileout)


def freeze_pan_for_all_sheet():
    for name in wb_out.sheetnames:
        wb_out[name].freeze_panes = 'D2'
    wb_out.save(fileout)


if __name__ == '__main__':
    category = get_categoty()
    # print(category)
    others = get_others()
    # print(others)
    gen_sheet()
    generate_cate_sheet()
    generate_top_sheet()
    generate_sl_sheet()
    freeze_pan_for_all_sheet()
