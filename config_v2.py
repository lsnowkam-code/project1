# config_v2.py - Обновленная конфигурация с поддержкой семантического поиска
"""
Расширенная версия config.py с поддержкой:
1. Ключевых слов для поиска колонок (вместо жестких индексов)
2. Нечеткого совпадения при поиске показателей
3. Динамического поиска года в шапке таблицы
"""

import csv
import re
import pandas as pd
from pathlib import Path
from typing import Dict, Tuple, Optional, List

# === Константы ===
EMPTY_CELL_MARKER = "—"
OKVED_CODE_COLUMN_INDEX = 0
OKVED_NAME_COLUMN_INDEX = 1


# === Утилиты ===
def _normalize_text(text: str) -> str:
    """Нормализует текст для сравнения"""
    return re.sub(r'\s+', ' ', text).strip().lower()


def _read_csv_robustly(filepath, header_row=0):
    """Читает CSV с несколькими попытками кодировки"""
    encodings = ['utf-8-sig', 'windows-1251', 'cp1251', 'utf-8', 'latin1']
    for encoding in encodings:
        try:
            print(f"Попытка чтения {filepath} с кодировкой {encoding}")
            df = pd.read_csv(filepath, sep=';', encoding=encoding, header=header_row)
            print(f"✅ Успешно прочитано с кодировкой {encoding}")
            if df.shape[1] == 1:
                print(f"⚠️ ВНИМАНИЕ: CSV-файл прочитан как один столбец. Возможно, проблема с разделителями.")
            return df
        except UnicodeDecodeError:
            pass
        except Exception:
            pass
    raise ValueError("❌ Не удалось прочитать CSV-файл.")


# === Новая функция: Загрузка маппингов с ключевыми словами ===
def load_column_mapping_v2(filepath) -> Tuple[Dict[str, str], Dict[str, Tuple[str, str]], 
                                               Dict[str, str], Dict, Dict[str, Dict[str, str]]]:
    """
    Загружает column_mapping_v2.csv с ключевыми словами.
    
    Returns:
        - word_to_indicator: назв ание показателя → код индикатора
        - indicator_to_excel: код индикатора → (колонка 2022, колонка 2023)
        - indicator_to_file: код индикатора → файл Excel
        - file_word_to_indicator: (файл, название) → код индикатора
        - indicator_keywords: код индикатора → {year_2022: [...], year_2023: [...]}
    """
    word_to_indicator = {}
    indicator_to_excel = {}
    indicator_to_file = {}
    file_word_to_indicator = {}
    indicator_keywords = {}

    try:
        with open(filepath, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=';')
            headers = next(reader)  # Пропускаем заголовок
            
            # Определяем индексы колонок
            col_idx = {name: idx for idx, name in enumerate(headers)}
            
            for row in reader:
                if len(row) < 3:
                    continue
                
                excel_file = str(row[col_idx.get('Excel файл', 0)]).strip()
                word_name = str(row[col_idx.get('Название показателя', 1)]).strip()
                indicator = str(row[col_idx.get('Код показателя', 2)]).strip()
                col_22 = str(row[col_idx.get('Excel колонка 2022', 3)]).strip() if len(row) > 3 else ''
                col_23 = str(row[col_idx.get('Excel колонка 2023', 4)]).strip() if len(row) > 4 else ''
                
                # Новые поля: ключевые слова
                kw_2022 = str(row[col_idx.get('Ключевое слово 2022', 5)]).strip() if len(row) > 5 else ''
                kw_2023 = str(row[col_idx.get('Ключевое слово 2023', 6)]).strip() if len(row) > 6 else ''
                
                if col_22.lower() == 'nan':
                    col_22 = ''
                if col_23.lower() == 'nan':
                    col_23 = ''
                
                word_to_indicator[word_name] = indicator
                indicator_to_excel[indicator] = (col_22, col_23)
                indicator_to_file[indicator] = excel_file
                file_word_to_indicator[(excel_file, _normalize_text(word_name))] = indicator
                
                # Парсим ключевые слова из кавычек
                keywords_2022 = [kw.strip('"').strip() for kw in kw_2022.split(',') if kw.strip()]
                keywords_2023 = [kw.strip('"').strip() for kw in kw_2023.split(',') if kw.strip()]
                
                indicator_keywords[indicator] = {
                    '2022': keywords_2022,
                    '2023': keywords_2023
                }
    
    except FileNotFoundError:
        print(f"⚠️ Файл {filepath} не найден, используем fallback на column_mapping.csv")
        # Fallback на старую версию
        from config import load_column_mapping
        word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator = load_column_mapping(filepath)
        indicator_keywords = {k: {'2022': [], '2023': []} for k in indicator_to_file.keys()}
    
    return word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator, indicator_keywords


# === Оригинальные функции (для совместимости) ===
def load_okved_map(filepath):
    print(f"Загрузка справочника ОКВЭД из: {filepath}")
    df = _read_csv_robustly(filepath, header_row=None)
    df = df.iloc[:, :2]
    df.columns = ['Код', 'Наименование']
    df = df.dropna(subset=['Код'])
    df['Код'] = df['Код'].astype(str).str.strip()
    df['Наименование'] = df['Наименование'].astype(str).str.strip()
    df['Наименование_норм'] = df['Наименование'].apply(_normalize_text)
    okved_to_name = dict(zip(df['Код'], df['Наименование']))
    name_to_okved_cleaned = dict(zip(df['Наименование_норм'], df['Код']))
    print(f"📘 Загружено {len(df)} записей ОКВЭД.")
    return okved_to_name, name_to_okved_cleaned


def load_table_source_map(filepath):
    print(f"Загрузка сопоставления таблиц из: {filepath}")
    df = _read_csv_robustly(filepath, header_row=None)
    first_row = [str(x).strip().lower() for x in df.iloc[0].tolist()]
    if "таблица" in first_row and "файл" in first_row:
        df = _read_csv_robustly(filepath, header_row=0)
    else:
        df.columns = ["Таблица", "Файл"]
    mapping_dict = {}
    for _, row in df.iterrows():
        raw_name = str(row["Таблица"]).strip()
        raw_src = str(row["Файл"]).strip()
        norm_key = _normalize_text(raw_name)
        mapping_dict[norm_key] = raw_src
    print(f"📘 Загружено сопоставлений: {len(mapping_dict)}")
    return mapping_dict
