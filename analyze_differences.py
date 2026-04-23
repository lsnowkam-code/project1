#!/usr/bin/env python3
"""Анализирует различия между сгенерированным и эталонным документом"""

import pandas as pd
from docx import Document
from collections import defaultdict

# Сравниваем документы
gen_doc = Document('/workspaces/project/output/Бюллетень_17.2.8_ГОТОВЫЙ.docx')
ref_doc = Document('/workspaces/project/output/как должно быть!.docx')

print("📊 АНАЛИЗ РАЗЛИЧИЙ ПО ТАБЛИЦАМ")
print("=" * 80)

# Группируем различия по таблицам
differences_by_table = defaultdict(list)

# Сравниваем таблицы
for table_idx in range(min(len(gen_doc.tables), len(ref_doc.tables))):
    gen_table = gen_doc.tables[table_idx]
    ref_table = ref_doc.tables[table_idx]

    table_differences = 0
    table_total_cells = 0

    for row_idx in range(min(len(gen_table.rows), len(ref_table.rows))):
        gen_row = gen_table.rows[row_idx]
        ref_row = ref_table.rows[row_idx]

        for col_idx in range(min(len(gen_row.cells), len(ref_row.cells))):
            gen_cell = gen_row.cells[col_idx].text.strip()
            ref_cell = ref_row.cells[col_idx].text.strip()

            table_total_cells += 1

            if gen_cell != ref_cell:
                table_differences += 1
                # Сохраняем только первые несколько отличий для каждой таблицы
                if table_differences <= 5:
                    differences_by_table[table_idx + 1].append({
                        'row': row_idx,
                        'col': col_idx,
                        'gen': gen_cell[:50],
                        'ref': ref_cell[:50]
                    })

    # Выводим статистику для каждой таблицы
    if table_differences > 0:
        print(f"\nТаблица {table_idx + 1}:")
        print(f"  Несовпадений: {table_differences}/{table_total_cells} ({table_differences/table_total_cells*100:.1f}%)")
        for diff in differences_by_table[table_idx + 1]:
            print(f"    Ячейка [{diff['row']},{diff['col']}]: сгеним={diff['gen'][:30]} vs эталон={diff['ref'][:30]}")

print("\n" + "=" * 80)
print(f"Всего таблиц с различиями: {len(differences_by_table)}")
print(f"Всего ячеек в таблицах с отличиями: {sum(len(diffs) for diffs in differences_by_table.values())}")
