#!/usr/bin/env bash
rm -f reduced.csv
python3 gen_finding_from_bao.py > reduced.csv
python3 gen_pyweb.py > RESULT.html