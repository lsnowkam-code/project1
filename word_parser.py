# word_parser.py
import re
from docx import Document

def normalize_text(text):
    """Приводит текст к нижнему регистру, убирает лишние пробелы и спецсимволы"""
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'\s+', ' ', text).strip()
    # Убираем скобки и содержимое в них для упрощения поиска (опционально)
    text = re.sub(r'\([^)]*\)', '', text).strip()
    # Убираем точки, запятые и другие пунктуационные знаки
    text = re.sub(r'[.,;:!?]', '', text).strip()
    return text

def parse_word_template(file_path):
    """
    Извлекает структуру таблиц из Word шаблона.
    Возвращает список словарей:
    [
      {
        'table_id': 1,
        'title': 'Валюта баланса',
        'years': ['2022', '2023'],
        'rows': [
            {'indicator_raw': 'Валюта баланса', 'indicator_norm': 'валюта баланса'},
            {'indicator_raw': 'Внеоборотные активы', 'indicator_norm': 'внеоборотные активы'},
            ...
        ]
      },
      ...
    ]
    """
    doc = Document(file_path)
    tables_data = []
    
    # Собираем заголовки из параграфов
    titles = {}
    for para in doc.paragraphs:
        text = para.text.strip()
        # Ищем заголовки вида "1. НАЗВАНИЕ", "2. НАЗВАНИЕ" и т.д.
        match = re.match(r'^(\d+)\.\s+(.+)', text)
        if match and not text.startswith('Продолжение'):
            table_num = int(match.group(1))
            table_name = match.group(2).strip()
            titles[table_num] = f"{table_num}. {table_name}"
    
    print(f"Найдено заголовков: {len(titles)}")
    
    # Обрабатываем таблицы
    tables_data = []
    for table_idx, table in enumerate(doc.tables):
        # Пропускаем пустые таблицы
        if len(table.rows) < 2:
            continue
        
        # Определяем номер таблицы (начиная с 1)
        table_num = table_idx + 1
        
        # Ищем соответствующий заголовок
        table_title = titles.get(table_num, f"Таблица №{table_num}")
        
        # 2. Поиск колонок с годами
        years_cols = {}
        header_row_idx = 0
        
        # Ищем строку с годами
        for r_idx, row in enumerate(table.rows[:3]):  # Смотрим первые 3 строки
            for c_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                if "2022" in text:
                    years_cols["2022"] = c_idx
                if "2023" in text:
                    years_cols["2023"] = c_idx
            if years_cols:
                header_row_idx = r_idx
                break
        
        if not years_cols:
            print(f"Пропускаем таблицу {table_num} - не найдены года")
            continue  # Таблица без лет
        
        # 3. Сбор показателей (строк данных)
        data_start_row = header_row_idx + 1
        rows_data = []
        
        for r_idx in range(data_start_row, len(table.rows)):
            row = table.rows[r_idx]
                
                # Извлекаем OKVED код из тегов в строке
                okved_code = None
                for cell in row.cells:
                    text = cell.text.strip()
                    # Ищем теги вида {{OKVED_КОД_ПОКАЗАТЕЛЬ_ГОД}}
                    import re
                    tag_match = re.search(r'{{OKVED_([^_]+)', text)
                    if tag_match:
                        okved_code = tag_match.group(1)
                        break
                
                if not okved_code:
                    continue  # Пропускаем строки без OKVED тегов
                
                rows_data.append({
                    "indicator_raw": okved_code,
                    "indicator_norm": normalize_text(okved_code),
                    "row_index": r_idx,
                    "target_cols": years_cols
                })
            "table_index": table_idx,
            "table_id": table_num,
            "title": table_title,
            "raw_header": table_title,
            "years_found": years_cols,
            "rows": rows_data
        })
    
    return tables_data

    return tables_data