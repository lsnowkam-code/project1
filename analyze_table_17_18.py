from docx import Document

generated = Document('/workspaces/project/output/Бюллетень_17.2.8_ГОТОВЫЙ.docx')
reference = Document('/workspaces/project/output/как должно быть!.docx')

# Таблица 17 (индекс 16)
print("=" * 80)
print("ТАБЛИЦА 17 - ДЕТАЛЬНО")
print("=" * 80)

table_gen = generated.tables[16]
table_ref = reference.tables[16]

print(f"Сгенерирована: {len(table_gen.rows)} строк x {len(table_gen.rows[0].cells)} столбцов")
print(f"Эталон:       {len(table_ref.rows)} строк x {len(table_ref.rows[0].cells)} столбцов")
print()

# Первые 5 строк
print("Первые 2 строки, столбцы 0-3 (сгенерирована):")
for r_idx in range(min(2, len(table_gen.rows))):
    for c_idx in range(min(4, len(table_gen.rows[0].cells))):
        text = table_gen.rows[r_idx].cells[c_idx].text[:30]
        print(f"  [{r_idx},{c_idx}]: {text}")

print()
print("Первые 2 строки, столбцы 0-3 (эталон):")
for r_idx in range(min(2, len(table_ref.rows))):
    for c_idx in range(min(4, len(table_ref.rows[0].cells))):
        text = table_ref.rows[r_idx].cells[c_idx].text[:30]
        print(f"  [{r_idx},{c_idx}]: {text}")

print()
print("=" * 80)
print("ТАБЛИЦА 18 - ДЕТАЛЬНО")
print("=" * 80)

table_gen = generated.tables[17]
table_ref = reference.tables[17]

print(f"Сгенерирована: {len(table_gen.rows)} строк x {len(table_gen.rows[0].cells)} столбцов")
print(f"Эталон:       {len(table_ref.rows)} строк x {len(table_ref.rows[0].cells)} столбцов")
print()

# Посмотрим заголовки
print("Заголовок (строка 0) сгенерирована:")
for idx, cell in enumerate(table_gen.rows[0].cells[:8]):
    print(f"  Столбец {idx}: {cell.text[:40]}")

print()
print("Заголовок (строка 0) эталон:")
for idx, cell in enumerate(table_ref.rows[0].cells[:8]):
    print(f"  Столбец {idx}: {cell.text[:40]}")

# Посчитаем заполненные ячейки по заголовкам
print()
print("Анализ сгенерированной таблицы 18:")
for col_idx in range(len(table_gen.rows[0].cells)):
    filled = sum(1 for row in table_gen.rows[1:] if table_gen.rows[row.index if hasattr(row, 'index') else table_gen.rows.index(row)].cells[col_idx].text.strip() and 
                 table_gen.rows[table_gen.rows.index(row) if hasattr(row, 'index') else len(table_gen.rows.index(row))].cells[col_idx].text.strip() != '—')
    if col_idx < 8:
        total = len(table_gen.rows) - 1
        print(f"  Столбец {col_idx}: {filled}/{total}")

