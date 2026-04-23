#!/usr/bin/env python3
"""Сравнение двух DOCX по таблицам с возможностью исключить таблицу по номеру.
Работает без python-docx (использует zip/xml стандартной библиотеки).
"""

import argparse
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def extract_tables(docx_path: str):
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))

    tables = []
    for tbl in root.findall('.//w:tbl', NS):
        rows = []
        for tr in tbl.findall('./w:tr', NS):
            cols = []
            for tc in tr.findall('./w:tc', NS):
                txt = ''.join((t.text or '') for t in tc.findall('.//w:t', NS)).strip()
                cols.append(txt)
            rows.append(cols)
        tables.append(rows)
    return tables


def compare_tables(gen_tables, ref_tables, skip_table=None):
    by_table = Counter()
    total_diff = 0
    matched_tables = 0
    checked_tables = 0

    max_tables = max(len(gen_tables), len(ref_tables))
    for idx in range(max_tables):
        table_num = idx + 1
        if skip_table is not None and table_num == skip_table:
            continue

        checked_tables += 1
        gtbl = gen_tables[idx] if idx < len(gen_tables) else []
        rtbl = ref_tables[idx] if idx < len(ref_tables) else []

        table_diff = 0
        for row_idx in range(max(len(gtbl), len(rtbl))):
            grow = gtbl[row_idx] if row_idx < len(gtbl) else []
            rrow = rtbl[row_idx] if row_idx < len(rtbl) else []

            for col_idx in range(max(len(grow), len(rrow))):
                gval = grow[col_idx] if col_idx < len(grow) else ''
                rval = rrow[col_idx] if col_idx < len(rrow) else ''
                if gval != rval:
                    table_diff += 1

        if table_diff == 0:
            matched_tables += 1
        else:
            by_table[table_num] = table_diff
            total_diff += table_diff

    return total_diff, by_table, matched_tables, checked_tables


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("generated")
    parser.add_argument("reference")
    parser.add_argument("--skip-table", type=int, default=2)
    args = parser.parse_args()

    gen_tables = extract_tables(args.generated)
    ref_tables = extract_tables(args.reference)

    total_diff, by_table, matched, checked = compare_tables(
        gen_tables, ref_tables, skip_table=args.skip_table
    )

    print(f"tables: generated={len(gen_tables)}, reference={len(ref_tables)}")
    print(f"skip_table: {args.skip_table}")
    print(f"total_diff: {total_diff}")
    print(f"tables_matched: {matched}/{checked}")
    print("diff_by_table:")
    for table_num, diff_count in by_table.most_common():
        print(f"  table {table_num}: {diff_count}")


if __name__ == "__main__":
    main()
