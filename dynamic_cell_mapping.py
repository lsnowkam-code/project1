"""
Динамическое сканирование и построение маппинга для каждой ячейки таблицы Word
Основано на table_source_data_mapping.csv и структуре Excel источников
"""

from docx import Document
from pathlib import Path
import pandas as pd
import csv
from difflib import SequenceMatcher

def similar(a, b):
    """Вычисляет похожесть двух строк"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def load_table_source_mapping():
    """Загружает маппинг таблиц"""
    mapping = {}
    mapping_file = Path('/workspaces/project/input/mappings/table_source_data_mapping.csv')
    
    with open(mapping_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            table_name = row['Таблица'].strip()
            source = row['Источник'].strip()
            mapping[table_name] = source
    
    return mapping

def extract_table_header(table, num_rows=5):
    """Извлекает заголовок таблицы"""
    header_text = []
    for i in range(min(num_rows, len(table.rows))):
        row_text = ' '.join([cell.text.strip() for cell in table.rows[i].cells if cell.text.strip()])
        if row_text:
            header_text.append(row_text)
    return ' '.join(header_text)

def find_matching_source(table_header, mapping):
    """Находит совпадающий источник для таблицы"""
    best_match = None
    best_score = 0.0
    
    for table_name, source in mapping.items():
        score = similar(table_header, table_name)
        if score > best_score:
            best_score = score
            best_match = (table_name, source)
    
    return best_match if best_score > 0.5 else None

def analyze_table_structure(table):
    """Анализирует структуру таблицы"""
    analysis = {
        'rows': len(table.rows),
        'cols': len(table.rows[0].cells) if table.rows else 0,
        'header_rows': 0,
        'data_rows': 0,
        'empty_rows': 0
    }
    
    # Определяем количество строк с заголовками
    for i, row in enumerate(table.rows):
        if i >= 5:  # Предполагаем максимум 5 строк заголовков
            break
        text = ' '.join([cell.text.strip() for cell in row.cells])
        if text and not any(cell.text.strip().replace(',', '').replace('.', '').isdigit() 
                           for cell in row.cells if cell.text.strip()):
            analysis['header_rows'] = i + 1
        else:
            break
    
    # Считаем строки с данными и пустые строки
    for i in range(analysis['header_rows'], len(table.rows)):
        row_has_data = False
        for cell in table.rows[i].cells:
            text = cell.text.strip()
            if text and text != '—':
                row_has_data = True
                break
        
        if row_has_data:
            analysis['data_rows'] += 1
        else:
            analysis['empty_rows'] += 1
    
    return analysis

def create_cell_mapping(doc_path, output_path):
    """Создаёт полный маппинг для каждой ячейки"""
    
    doc = Document(doc_path)
    mapping = load_table_source_mapping()
    
    print("=" * 100)
    print("ДИНАМИЧЕСКОЕ СКАНИРОВАНИЕ И СОЗДАНИЕ МАППИНГА ЯЧЕЕК")
    print("=" * 100)
    print()
    
    all_mappings = []
    
    for table_idx, table in enumerate(doc.tables):
        print(f"📊 Таблица {table_idx + 1}:")
        
        # Извлекаем заголовок
        header = extract_table_header(table)
        print(f"   Заголовок: {header[:80]}...")
        
        # Ищем источник
        source_match = find_matching_source(header, mapping)
        if source_match:
            table_name, source = source_match
            print(f"   ✅ Источник найден: {source}")
        else:
            print(f"   ⚠️  Источник не найден")
            source = "UNKNOWN"
        
        # Анализируем структуру
        structure = analyze_table_structure(table)
        print(f"   Структура: {structure['rows']} строк × {structure['cols']} столбцов")
        print(f"   Заголовки: {structure['header_rows']} строк, Данные: {structure['data_rows']} строк")
        print()
        
        # Создаём маппинг для каждой ячейки
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                
                # Определяем тип ячейки
                if row_idx < structure['header_rows']:
                    cell_type = 'HEADER'
                elif col_idx == 0:
                    cell_type = 'OKVED_CODE'
                elif col_idx == 1:
                    cell_type = 'OKVED_NAME'
                else:
                    if cell_text and cell_text != '—':
                        cell_type = 'DATA'
                    else:
                        cell_type = 'EMPTY'
                
                # Добавляем в маппинг
                all_mappings.append({
                    'table_num': table_idx + 1,
                    'table_name': table_name if source_match else 'UNKNOWN',
                    'source': source,
                    'row': row_idx,
                    'col': col_idx,
                    'cell_type': cell_type,
                    'cell_value': cell_text[:100],
                    'structure_header_rows': structure['header_rows'],
                    'structure_data_rows': structure['data_rows']
                })
    
    # Сохраняем маппинг
    df = pd.DataFrame(all_mappings)
    df.to_csv(output_path, encoding='utf-8', index=False)
    
    print("=" * 100)
    print(f"✅ Маппинг сохранён: {output_path}")
    print(f"   Всего ячеек отображено: {len(all_mappings)}")
    print("=" * 100)
    
    return df

if __name__ == '__main__':
    doc_path = '/workspaces/project/output/Бюллетень_17.2.8_ГОТОВЫЙ.docx'
    output_path = '/workspaces/project/output/динамический_маппинг_ячеек.csv'
    
    df = create_cell_mapping(doc_path, output_path)
    
    # Статистика
    print()
    print("📊 СТАТИСТИКА МАППИНГА:")
    print()
    print(f"По типам ячеек:")
    print(df['cell_type'].value_counts())
    print()
    print(f"По таблицам:")
    print(df['table_num'].value_counts().sort_index())
