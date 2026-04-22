from docx import Document

generated = Document('/workspaces/project/output/Бюллетень_ПРОВЕРКА_ПРАВИЛЬНО.docx')
reference = Document('/workspaces/project/output/как должно быть!.docx')

def count_filled_cells(table):
    """Подсчёт заполненных ячеек"""
    total = 0
    filled = 0
    for row in table.rows:
        for cell in row.cells:
            text = cell.text.strip()
            total += 1
            if text and text != '—':
                filled += 1
    return filled, total

print("=" * 100)
print("ПОЛНОЕ СРАВНЕНИЕ ВСЕХ 19 ТАБЛИЦ")
print("=" * 100)
print()

total_match = 0
total_tables = min(len(generated.tables), len(reference.tables))

for idx in range(total_tables):
    gen_table = generated.tables[idx]
    ref_table = reference.tables[idx]
    
    gen_filled, gen_total = count_filled_cells(gen_table)
    ref_filled, ref_total = count_filled_cells(ref_table)
    
    gen_pct = (gen_filled / gen_total * 100) if gen_total > 0 else 0
    ref_pct = (ref_filled / ref_total * 100) if ref_total > 0 else 0
    
    match = gen_pct == ref_pct
    if match:
        total_match += 1
        status = "✅"
    else:
        status = "❌"
    
    print(f"Таблица {idx + 1:2d}: {status} {gen_pct:5.1f}% (сгенерирована) vs {ref_pct:5.1f}% (эталон)", end="")
    if not match:
        print(f" [разница: {abs(gen_pct - ref_pct):+.1f}%]")
    else:
        print()

print()
print("=" * 100)
print(f"РЕЗУЛЬТАТ: {total_match}/{total_tables} таблиц совпадают ({total_match*100/total_tables:.1f}%)")
print("=" * 100)

if total_match == total_tables:
    print("🎉 ВСЕ ТАБЛИЦЫ СОВПАДАЮТ С ЭТАЛОНОМ!")
else:
    print(f"⚠️ {total_tables - total_match} таблиц не совпадают")
