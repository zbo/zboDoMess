import csv
import comlib

new_added = []
meta_all_codes = []
repeat = []

with open("raw.txt", "r") as fo:
    all_lines = fo.readlines()
    for line in all_lines:
        if 7 == len(line):
            code = line[0:6]
            if code[0] == '0' or code[0] == '3':
                new_added.append(code + '.sz')
            if code[0] == '6':
                new_added.append(code + '.sh')

meta_all_codes = comlib.get_orign_codes_from_xl()
for item in new_added:
    if item not in meta_all_codes:
        print(item)
    else:
        repeat.append(item)
print('repeated include:{0}'.format(repeat))
