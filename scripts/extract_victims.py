"""Extract county-level abused child/youth counts (2015-2024) from the raw xlsx.

Usage: python scripts/extract_victims.py data/raw/abused_children_by_county_2015_2024.xlsx data/abused_children_by_county_2015_2024.csv
"""
import csv
import sys

import openpyxl


def main(src, dst):
    wb = openpyxl.load_workbook(src, data_only=True)
    years = sorted(int(n) for n in wb.sheetnames if n.isdigit())
    out = []
    for y in years:
        rows = list(wb[str(y)].iter_rows(values_only=True))
        for i, r in enumerate(rows):
            if r[0] and r[1] == '計' and isinstance(r[3], (int, float)):
                name = str(r[0]).replace('　', '').replace(' ', '').strip()
                indig = next((rr[3] for rr in rows[i + 1:i + 3] if rr[1] == '原住民'), None)
                out.append([name, y, y - 1911, r[3], r[4], r[5], indig])
    with open(dst, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['county', 'year', 'roc_year', 'total', 'male', 'female', 'indigenous'])
        w.writerows(out)
    print(f'{len(out)} rows -> {dst}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
