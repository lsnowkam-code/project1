from docx import Document
from pathlib import Path

def count_filled_cells(table):
    """Подсчёт заполненных ячеек в таблице"""
    total = 0
    filled = 0
    for row in table.rows:
        for cell in row.cells:
            text = cell.text.strip()
            total += 1
            # Считаем заполненной если не пусто и не только дефис
            if text and text != '—':
                filled += 1
    return filled, total

# Загружаем оба документа
generated = Document('/workspaces/project/output/Бюллетень_17.2.8_ГОТОВЫЙ.docx')
reference = Document('/workspaces/project/output/как должно быть!.docx')

print("=" * 80)
print("ДЕТАЛЬНОЕ СРАВНЕНИЕ ЗАПОЛНЕНИЯ ТАБЛИЦ")
print("=" * 80)
print()

print(f"Таблиц в сгенерированном: {len(generated.tables)}")
print(f"Таблиц в эталоне: {len(reference.tables)}")
print()

# Анализируем каждую таблицу
for idx in range(min(len(generated.tables), len(reference.tables))):
    gen_table = generated.tables[idx]
    ref_table = reference.tables[idx]
    
    gen_filled, gen_total = count_filled_cells(gen_table)
    ref_filled, ref_total = count_filled_cells(ref_table)
    
    gen_pct = (gen_filled / gen_total * 100) if gen_total > 0 else 0
    ref_pct = (ref_filled / ref_total * 100) if ref_total > 0 else 0
    
    # Проверяем совпадение
    match = "✅ СОВПАДАЕТ" if gen_pct == ref_pct else "❌ ОТЛИЧАЕТСЯ"
    
    print(f"Таблица {idx + 1}:")
    print(f"  Сгенерирована:    {gen_filled:5d}/{gen_total:5d} ({gen_pct:5.1f}%) - {gen_table.rows[0].cells[0].text[:30]}")
    print(f"  Эталон:           {ref_filled:5d}/{ref_total:5d} ({ref_pct:5.1f}%) - {ref_table.rows[0].cells[0].text[:30]}")
    print(f"  Статус: {match}")
    if gen_pct != ref_pct:
        print(f"  Разница: {abs(gen_pct - ref_pct):.1f}%")
    print()

print("=" * 80)
