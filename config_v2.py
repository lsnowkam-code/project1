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
            # Если нет заголовка, явно задаём имена столбцов
            if header_row is None:
                df.columns = [f"col_{i}" for i in range(df.shape[1])]
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
            lines = f.readlines()
            
            # Пропускаем заголовок
            if lines:
                lines = lines[1:]
            
            for line in lines:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                
                # Убираем кавычки в начале и конце строки, если они есть
                if line.startswith('"') and line.endswith('"'):
                    line = line[1:-1]
                
                # Разделяем по точке с запятой
                row = line.split(';')
                
                if len(row) < 3:
                    continue
                
                excel_file = str(row[0]).strip()
                word_name = str(row[1]).strip()
                indicator = str(row[2]).strip()
                col_22 = str(row[3]).strip() if len(row) > 3 else ''
                col_23 = str(row[4]).strip() if len(row) > 4 else ''
                
                # Новые поля: ключевые слова (если есть)
                kw_2022 = str(row[5]).strip() if len(row) > 5 else ''
                kw_2023 = str(row[6]).strip() if len(row) > 6 else ''
                
                if col_22.lower() == 'nan':
                    col_22 = ''
                if col_23.lower() == 'nan':
                    col_23 = ''
                
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
    if "таблица" in first_row and ("файл" in first_row or "источник" in first_row):
        df = _read_csv_robustly(filepath, header_row=0)
    else:
        df.columns = ["Таблица", "Файл"]
    mapping_dict = {}
    file_column = "Файл" if "Файл" in df.columns else "Источник" if "Источник" in df.columns else df.columns[1]
    for _, row in df.iterrows():
        raw_name = str(row["Таблица"]).strip()
        raw_src = str(row[file_column]).strip()
        norm_key = _normalize_text(raw_name)
        mapping_dict[norm_key] = raw_src
    print(f"📘 Загружено сопоставлений: {len(mapping_dict)}")
    return mapping_dict


def _detect_year_by_text(text: str) -> str:
    if not isinstance(text, str):
        return None
    s = text.lower()
    if 'преды' in s or '2022' in s:
        return '22'
    if 'отчет' in s or '2023' in s or 'текущ' in s:
        return '23'
    return None


def _find_excel_header_row(df: pd.DataFrame) -> int:
    for idx in range(min(20, len(df))):
        row = df.iloc[idx, :10].astype(str).fillna('').str.lower().tolist()
        joined = ' '.join(row)
        if 'код' in joined and 'наименование' in joined:
            return idx
    return None


def _slugify_indicator_code(text: str, prefix: str = None) -> str:
    if not text:
        return ''
    normalized = _normalize_text(text)
    code = re.sub(r'[^0-9a-zа-я]+', '_', normalized)
    code = re.sub(r'_+', '_', code).strip('_')
    if prefix:
        prefix_norm = re.sub(r'[^0-9a-zа-я]+', '_', _normalize_text(prefix))
        prefix_norm = re.sub(r'_+', '_', prefix_norm).strip('_')
        return f"{prefix_norm}_{code}" if code else prefix_norm
    return code


def _is_connective_header(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return _normalize_text(text) in {
        'в том числе',
        'в том числе:'
    }


def _infer_mapping_from_excel(excel_path: Path) -> list:
    df = pd.read_excel(excel_path, header=None)
    header_row = _find_excel_header_row(df)
    if header_row is None:
        print(f"⚠️ Не найден заголовок с 'Код' и 'Наименование' в {excel_path.name}")
        return []

    headers = df.iloc[header_row].astype(str).fillna('').tolist()
    subheaders = df.iloc[header_row + 1].astype(str).fillna('').tolist() if header_row + 1 < len(df) else [''] * len(headers)

    current_base = ''
    current_indicator_name = None
    groups = {}
    for col_idx in range(2, len(headers)):
        raw_header = str(headers[col_idx]).strip()
        suffix = str(subheaders[col_idx]).strip()
        year = _detect_year_by_text(suffix)

        if raw_header:
            if _is_connective_header(raw_header):
                if not current_base or not suffix:
                    continue
                indicator_name = f"{current_base} {raw_header} {suffix}".strip()
            else:
                current_base = raw_header
                if suffix and not year:
                    indicator_name = f"{current_base} {suffix}".strip()
                else:
                    indicator_name = current_base
        elif suffix:
            if not current_base:
                continue
            if year:
                indicator_name = current_indicator_name or current_base
            else:
                indicator_name = f"{current_base} {suffix}".strip()
        else:
            if not current_indicator_name:
                continue
            indicator_name = current_indicator_name

        current_indicator_name = indicator_name
        group_key = _normalize_text(indicator_name)
        if group_key not in groups:
            groups[group_key] = {
                'name': indicator_name,
                'code': _slugify_indicator_code(indicator_name, excel_path.stem),
                'cols': {'22': None, '23': None},
                'keywords': [current_base]
            }

        if year:
            groups[group_key]['cols'][year] = str(col_idx + 1)
        else:
            if groups[group_key]['cols']['22'] is None:
                groups[group_key]['cols']['22'] = str(col_idx + 1)
            else:
                groups[group_key]['cols']['23'] = str(col_idx + 1)

    rows = []
    for group in groups.values():
        col_22 = group['cols']['22'] or ''
        col_23 = group['cols']['23'] or ''
        keyword_2022 = group['keywords'][0] if group['keywords'] else ''
        keyword_2023 = group['keywords'][0] if group['keywords'] else ''
        rows.append((excel_path.name, group['name'], group['code'], col_22, col_23, keyword_2022, keyword_2023))

    return rows


def build_column_mapping_v2_from_excel(excel_dir: Path, table_source_mapping: dict, output_path: Path):
    excel_dir = Path(excel_dir)
    output_path = Path(output_path)
    rows = []
    seen_codes = set()

    for filename in sorted(set(table_source_mapping.values())):
        excel_path = excel_dir / filename
        if not excel_path.exists():
            print(f"⚠️ Excel файл не найден: {filename}")
            continue
        inferred = _infer_mapping_from_excel(excel_path)
        for excel_file, name, code, col_22, col_23, kw22, kw23 in inferred:
            if code in seen_codes:
                suffix = 1
                base_code = code
                while f"{base_code}_{suffix}" in seen_codes:
                    suffix += 1
                code = f"{base_code}_{suffix}"
            seen_codes.add(code)
            rows.append((excel_file, name, code, col_22, col_23, kw22, kw23))

    if not rows:
        raise ValueError("Не удалось сгенерировать column_mapping_v2 из Excel.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            'Excel файл', 'Название показателя', 'Код показателя',
            'Excel колонка 2022', 'Excel колонка 2023',
            'Ключевое слово 2022', 'Ключевое слово 2023'
        ])
        for row in rows:
            writer.writerow(row)

    print(f"✅ Сгенерирован column_mapping_v2: {output_path} ({len(rows)} строк)")
