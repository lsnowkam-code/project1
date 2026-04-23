#!/usr/bin/env python3
"""Глубокая проверка DOCX (как ZIP/XML) на дубли тегов внутри строк таблиц.
Пример:
  python deep_docx_tag_audit.py output/Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx --tables 5 8 14 17
"""

import argparse
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
TAG_RE = re.compile(r"\{\{([^}]+)\}\}")


def load_tables(docx_path: str):
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return root.findall('.//w:tbl', NS)


def row_tags(row):
    tags = []
    for cell in row.findall('./w:tc', NS):
        txt = ''.join((t.text or '') for t in cell.findall('.//w:t', NS))
        tags.extend(TAG_RE.findall(txt))
    return tags


def audit_table(table):
    rows_with_duplicates = 0
    extra_duplicate_occurrences = 0
    duplicate_kinds = Counter()

    for row in table.findall('./w:tr', NS):
        tags = row_tags(row)
        if not tags:
            continue
        c = Counter(tags)
        duplicated = {k: v for k, v in c.items() if v > 1}
        if duplicated:
            rows_with_duplicates += 1
            extra_duplicate_occurrences += sum(v - 1 for v in duplicated.values())
            for tag, cnt in duplicated.items():
                duplicate_kinds[tag] += cnt

    return {
        "rows_with_duplicates": rows_with_duplicates,
        "extra_duplicate_occurrences": extra_duplicate_occurrences,
        "duplicate_tag_kinds": len(duplicate_kinds),
        "top_duplicate_tags": duplicate_kinds.most_common(10),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--tables", nargs="*", type=int, default=[5, 8, 14, 17])
    args = parser.parse_args()

    tables = load_tables(args.docx)
    print(f"docx={args.docx}")
    print(f"tables_total={len(tables)}")

    for table_num in args.tables:
        if table_num < 1 or table_num > len(tables):
            print(f"table {table_num}: out_of_range")
            continue

        result = audit_table(tables[table_num - 1])
        print(
            f"table {table_num}: rows_with_duplicates={result['rows_with_duplicates']}, "
            f"extra_duplicate_occurrences={result['extra_duplicate_occurrences']}, "
            f"duplicate_tag_kinds={result['duplicate_tag_kinds']}"
        )
        if result["top_duplicate_tags"]:
            print(f"  top_duplicate_tags={result['top_duplicate_tags'][:5]}")


if __name__ == "__main__":
    main()
