# smart_excel_finder.py
import pandas as pd
from rapidfuzz import process, fuzz

def normalize_excel_text(text):
    """Нормализация текста из Excel для сравнения"""
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    # Убираем лишние пробелы, скобки, точки в конце
    import re
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[\.]', '', text)
    return text

def find_value(excel_file_path, okved_normalized, indicator_normalized, year):
    """
    Ищет значение в Excel файле.
    :param excel_file_path: Путь к файлу .xlsx
    :param okved_normalized: Нормализованный OKVED код (из Word)
    :param indicator_normalized: Нормализованное имя показателя (из заголовка таблицы)
    :param year: Год ('2022' или '2023')
    :return: Найденное значение или None
    """
    try:
        df = pd.read_excel(excel_file_path, header=None)
        
        # 1. Найти колонку для показателя (в Row[5])
        indicator_col_index = None
        for col_idx, cell_val in enumerate(df.iloc[5]):  # Row[5] - показатели
            cell_str = normalize_excel_text(cell_val)
            if cell_str and (indicator_normalized in cell_str or cell_str in indicator_normalized):
                indicator_col_index = col_idx
                break
        
        if indicator_col_index is None:
            return None  # Не нашли колонку с показателем

        # 2. Найти суб-колонку для года (в Row[6], пары колонок для каждого показателя)
        year_col_index = None
        year_text = "на конец отчетного года" if year == "2023" else "на конец предыдущего года"
        
        # Показатели начинаются с колонки indicator_col_index, годы чередуются
        # Для каждого показателя: col, col+1 (предыдущий, отчетный)
        base_col = indicator_col_index
        if year == "2022":
            year_col_index = base_col  # предыдущий год
        else:
            year_col_index = base_col + 1  # отчетный год
        
        # 3. Найти строку по OKVED (в столбце 0, начиная с Row[8])
        okved_row_index = None
        for idx in range(8, len(df)):  # Начиная с Row[8] - данные
            val = df.iloc[idx, 0]
            if pd.notna(val):
                val_str = normalize_excel_text(str(val))
                if val_str == okved_normalized:  # Точное совпадение
                    okved_row_index = idx
                    print(f"Found exact OKVED match '{okved_normalized}' in row {idx}")
                    break
        
        if okved_row_index is None:
            print(f"No exact match for OKVED '{okved_normalized}'")
            return None  # Не нашли строку с OKVED

        # 4. Извлечь значение
        value = df.iloc[okved_row_index, year_col_index]
        return value

    except Exception as e:
        print(f"Ошибка при чтении {excel_file_path}: {e}")
        return None