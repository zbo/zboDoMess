import csv

def gen_meta():
    with open('list.csv', 'r') as f:
        meta_csv_reader = csv.reader(f)
        meta_all_codes = []
        meta_all_urls = []
        for row in meta_csv_reader:
            if row in meta_all_codes:
                print(row)
                continue

if __name__ == '__main__':
    gen_meta()