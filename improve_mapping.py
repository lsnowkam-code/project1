"""
Улучшаем маппинг с адекватным определением источников
"""
from docx import Document
import pandas as pd

def get_full_header(table):
    """Get full header including first data rows if needed"""
    parts = []
    for i in range(min(5, len(table.rows))):
        for cell in table.rows[i].cells:
            text = cell.text.strip()
            if text:
                parts.append(text)
    return ' '.join(parts)

# Расширенный список источников
SOURCES_MAP = {
    '1': ('БАЛАНС', 'T23_000000_t01Ved14.xlsx'),
    '3': ('ВНЕОБОРОТНЫЕ АКТИВЫ', 'T23_000000_t03Ved14.xlsx'),
    '4': ('ОБОРОТНЫЕ АКТИВЫ', 'T23_000000_t10Ved14.xlsx'),
    '5': ('ОБОРАЧИВАЕМОСТЬ', 'T23_000000_t13Ved14.xlsx'),
    '6': ('КАПИТАЛ И РЕЗЕРВЫ', 'T23_000000_t19Ved14.xlsx'),
    '7': ('ДОЛГОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА', 'T23_000000_t20Ved14.xlsx'),
    '8': ('КРАТКОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА', 'T23_000000_t21Ved14.xlsx'),
}

# Анализ неизвестных таблиц
doc = Document('/workspaces/project/output/как должно быть!.docx')

unknown_tables = [6, 7, 12, 13, 14, 15, 16, 18]  # 0-based индексы

print("=" * 100)
print("АНАЛИЗ НЕИЗВЕСТНЫХ ТАБЛИЦ")
print("=" * 100)
print()

for table_idx in unknown_tables:
    table = doc.tables[table_idx]
    header = get_full_header(table)
    
    print(f"Таблица {table_idx + 1}:")
    print(f"  Размер: {len(table.rows)} × {len(table.rows[0].cells)}")
    print(f"  Заголовок: {header[:100]}...")
    
    # Пытаемся найти номер таблицы в заголовке
    import re
    match = re.search(r'(\d+)\.', header)
    if match:
        table_num = match.group(1)
        if table_num in SOURCES_MAP:
            name, source = SOURCES_MAP[table_num]
            print(f"  ✅ Найден номер: {table_num} -> {source}")
        else:
            print(f"  ❌ Номер {table_num} не в маппинге")
    else:
        print(f"  ❌ Номер не найден в заголовке")
    
    # Анализируем содержимое
    sample_data = []
    for i in range(min(5, len(table.rows))):
        for j in range(min(3, len(table.rows[0].cells))):
            text = table.rows[i].cells[j].text.strip()
            if text and text != '—':
                sample_data.append(text[:40])
    
    if sample_data:
        print(f"  Примеры: {', '.join(sample_data[:3])}")
    print()

print("=" * 100)
print("\n💡 РЕКОМЕНДАЦИИ:")
print("  Таблица 6: похоже на дополнительную таблицу")
print("  Таблица 7: похоже на дополнительную таблицу")
print("  Таблицы 12-16, 18: требуют ручного определения по доступным источникам")
print("=" * 100)
