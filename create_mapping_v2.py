"""
Create dynamic mapping for each cell based on table sources
"""
from docx import Document
import pandas as pd
import re

def get_table_header(table, rows=3):
    """Get table header text"""
    parts = []
    for i in range(min(rows, len(table.rows))):
        for cell in table.rows[i].cells:
            text = cell.text.strip()
            if text:
                parts.append(text)
    return ' '.join(parts)

def get_table_source(header_text):
    """Determine the Excel source based on header"""
    # Сопоставляем с известными источниками
    sources = {
        'БАЛАНС': ('БАЛАНС ОРГАНИЗАЦИЙ', 'T23_000000_t01Ved14.xlsx'),
        'ВНЕОБОРОТНЫЕ АКТИВЫ': ('ВНЕОБОРОТНЫЕ АКТИВЫ', 'T23_000000_t03Ved14.xlsx'),
        'ОБОРОТНЫЕ АКТИВЫ': ('ОБОРОТНЫЕ АКТИВЫ', 'T23_000000_t10Ved14.xlsx'),
        'ОБОРАЧИВАЕМОСТЬ': ('ОБОРАЧИВАЕМОСТЬ', 'T23_000000_t13Ved14.xlsx'),
        'КАПИТАЛ И РЕЗЕРВЫ': ('КАПИТАЛ И РЕЗЕРВЫ', 'T23_000000_t19Ved14.xlsx'),
        'ДОЛГОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА': ('ДОЛГОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА', 'T23_000000_t20Ved14.xlsx'),
        'КРАТКОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА': ('КРАТКОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА', 'T23_000000_t21Ved14.xlsx'),
    }
    
    header_upper = header_text.upper()
    
    for keyword, (name, source) in sources.items():
        if keyword in header_upper:
            return name, source
    
    return 'UNKNOWN', 'UNKNOWN'

print("=" * 100)
print("ПОСТРОЕНИЕ ДИНАМИЧЕСКОГО МАППИНГА ЯЧЕЕК")
print("=" * 100)
print()

doc = Document('/workspaces/project/output/как должно быть!.docx')

all_mappings = []

for table_idx, table in enumerate(doc.tables):
    header = get_table_header(table)
    table_name, source = get_table_source(header)
    
    print(f"📊 Таблица {table_idx + 1}: {table_name} ({source})")
    
    # Определяем структуру таблицы
    header_rows = 0
    for i in range(min(5, len(table.rows))):
        row_data = [cell.text.strip() for cell in table.rows[i].cells]
        # Проверяем есть ли числа
        has_data = any(
            any(c.isdigit() for c in cell.replace('.', '').replace(',', ''))
            for cell in row_data if cell and cell != '—'
        )
        if not has_data:
            header_rows = i + 1
        else:
            break
    
    print(f"   Размер: {len(table.rows)} строк × {len(table.rows[0].cells)} столбцов")
    print(f"   Заголовки: {header_rows} строк")
    
    # Создаём маппинг для каждой ячейки
    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            cell_text = cell.text.strip()
            
            # Определяем тип ячейки
            if row_idx < header_rows:
                cell_type = 'HEADER'
            elif col_idx == 0:
                cell_type = 'OKVED_CODE'
            elif col_idx == 1 and row_idx >= header_rows:
                cell_type = 'OKVED_NAME'
            elif not cell_text:
                cell_type = 'EMPTY'
            elif cell_text == '—':
                cell_type = 'EMPTY_MARK'
            else:
                cell_type = 'DATA'
            
            all_mappings.append({
                'table_num': table_idx + 1,
                'table_name': table_name,
                'excel_source': source,
                'row': row_idx,
                'col': col_idx,
                'is_header_row': row_idx < header_rows,
                'cell_type': cell_type,
                'value': cell_text[:100] if cell_text else ''
            })
    print()

# Сохраняем маппинг
df = pd.DataFrame(all_mappings)
output_file = '/workspaces/project/output/ДИНАМИЧЕСКИЙ_МАППИНГ_ЯЧЕЕК.csv'
df.to_csv(output_file, encoding='utf-8-sig', index=False)

print("=" * 100)
print(f"✅ Маппинг сохранён в: {output_file}")
print(f"\n📊 СТАТИСТИКА МАППИНГА:")
print(f"\nВсего ячеек: {len(all_mappings)}")
print(f"\nПо типам ячеек:")
for cell_type in sorted(df['cell_type'].unique()):
    count = (df['cell_type'] == cell_type).sum()
    pct = count / len(df) * 100
    print(f"  {cell_type:20} : {count:6} ({pct:5.1f}%)")

print(f"\nПо таблицам:")
for table_num in sorted(df['table_num'].unique()):
    subset = df[df['table_num'] == table_num]
    table_name = subset['table_name'].iloc[0]
    count = len(subset)
    print(f"  Таблица {table_num:2}: {table_name:50} - {count:5} ячеек")

print("\n" + "=" * 100)
