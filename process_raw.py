with open("raw.txt", "r") as fo:
    all_lines = fo.readlines()
    for line in all_lines:
        if 7 == len(line):
            code = line[0:6]
            if code[0] == '0' or code[0] == '3':
                print(code + '.sz')
            if code[0] == '6':
                print(code + '.sh')
