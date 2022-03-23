import csv


def gen_meta():
    with open('list.csv', 'r') as f:
        meta_csv_reader = csv.reader(f)
        meta_all_codes = []
        for row in meta_csv_reader:
            if row in meta_all_codes:
                print(row)
                continue
            meta_all_codes.append(row)


if __name__ == '__main__':
    gen_meta()
