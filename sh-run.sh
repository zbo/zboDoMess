#!/usr/bin/env bash
#python2.7 gen.py >README.md
rm -f ./history.xlsx
#rm -f ./list.csv
#python process_raw.py > list.csv
#python3 sina_batch.py
python ./history_netease.py
open .