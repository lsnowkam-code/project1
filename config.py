# confic.py
import csv
import re

import pandas as pd
from docx.table import _Cell, Table

# === Константы ===
EMPTY_CELL_MARKER = "—"
OKVED_CODE_COLUMN_INDEX = 0
OKVED_NAME_COLUMN_INDEX = 1


# === Утилиты ===
def _normalize_text(text: str) -> str:
    return re.sub(r'\s+', ' ', text).strip().lower()


def get_cleaned_cell_text(cell: _Cell) -> str:
    return ' '.join(p.text.replace('\n', ' ').strip() for p in cell.paragraphs).strip()


def _read_csv_robustly(filepath, header_row=0):
    encodings = ['utf-8-sig', 'windows-1251', 'cp1251', 'utf-8', 'latin1']
    for encoding in encodings:
        try:
            print(f"Попытка чтения {filepath} с кодировкой {encoding}")
            df = pd.read_csv(filepath, sep=';', encoding=encoding, header=header_row)
            print(f"✅ Успешно прочитано с кодировкой {encoding}")
            if df.shape[1] == 1:
                print(
                    f"⚠️ ВНИМАНИЕ: CSV-файл прочитан как один столбец. Возможно, проблема с разделителями или кавычками.")
                print(f"📋 Заголовок: {df.columns[0]}")
                print("💡 Проверь, что файл сохранён с разделителями ';' и без двойных кавычек вокруг всей строки.")
            return df
        except UnicodeDecodeError as e:
            print(f"❌ Ошибка кодировки {encoding}: {e}")
        except Exception as e:
            print(f"❌ Другая ошибка с {encoding}: {e}")
    raise ValueError("❌ Не удалось прочитать CSV-файл ни с одной кодировкой.")


# === Загрузка справочников ===
def load_okved_map(filepath):
    print(f"Загрузка справочника ОКВЭД из: {filepath}")
    df = _read_csv_robustly(filepath, header_row=None)
    # .iloc — это позиционный индексатор pandas.→ взять все строки.→ взять только первые два столбца (с индексами 0
    # и 1, правая граница не включается).
    # Полученная таблица перезаписывает df. Лишние колонки отбрасываются.
    df = df.iloc[:, :2]
    # -----

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


def load_column_mapping(filepath):
    """
    Загружает column_mapping.csv и возвращает:
    - word_to_indicator: название показателя → код индикатора
    - indicator_to_excel: код индикатора → (номер колонки 2022, номер колонки 2023)
    - indicator_to_file: код индикатора → имя Excel-файла
    """
    word_to_indicator = {}
    indicator_to_excel = {}
    indicator_to_file = {}

    with open(filepath, encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        headers = next(reader)  # Пропускаем заголовок

        for row in reader:
            if len(row) < 5:
                continue  # Пропускаем неполные строки

            excel_file = row[0].strip()
            word_name = row[1].strip()
            indicator = row[2].strip()
            col_22 = row[3].strip()
            col_23 = row[4].strip()

            word_to_indicator[word_name] = indicator
            indicator_to_excel[indicator] = (col_22, col_23)
            indicator_to_file[indicator] = excel_file

    return word_to_indicator, indicator_to_excel, indicator_to_file


# === Поиск названия таблицы ===
def get_table_name(table: Table, known_table_names):
    ignore_phrases = ['продолжение таблицы', 'таблица', 'график', 'рис.', 'рис', 'по всей форме', 'приложение', 'форма',
                      'лист', 'страница']
    prev_elem = table._element.getprevious()
    while prev_elem is not None:
        if prev_elem.tag.endswith('p'):
            text = (prev_elem.text or "").strip()
            if not text:
                prev_elem = prev_elem.getprevious()
                continue
            text_norm = _normalize_text(text)
            if any(phrase in text_norm for phrase in ignore_phrases):
                prev_elem = prev_elem.getprevious()
                continue
            for known_title in known_table_names:
                if text_norm in _normalize_text(known_title) or _normalize_text(known_title) in text_norm:
                    print(f"🔍 Заголовок найден: {known_title}")
                    return known_title
            print(f"⚠️ Заголовок не распознан: {text}")
            return text_norm
        prev_elem = prev_elem.getprevious()
    collected_texts = []
    for row in table.rows[:5]:
        for cell in row.cells[:3]:
            txt = get_cleaned_cell_text(cell)
            if txt:
                collected_texts.append(txt)
    combined = _normalize_text(" ".join(collected_texts))
    for known_title in known_table_names:
        if _normalize_text(known_title) in combined or combined in _normalize_text(known_title):
            print(f"🔍 Заголовок найден внутри таблицы: {known_title}")
            return known_title
    print("⚠️ Не удалось определить заголовок таблицы.")
    return ""


# === Загрузка Excel-данных ===
def get_excel_data(excel_path, okved_codes_set):
    try:
        df = pd.read_excel(excel_path, header=None, dtype=str)
    except Exception as e:
        print(f"❌ Ошибка чтения Excel: {e}")
        return {}

    header_map = {}
    header_row_index = -1

    # 🔍 Поиск строки заголовков (обычно строка с 'А' и '1')
    for i, row in df.head(15).iterrows():
        if len(row) > 2 and pd.notna(row.iloc[0]) and pd.notna(row.iloc[2]):
            if str(row.iloc[0]).strip().upper() == 'А' and str(row.iloc[2]).strip() == '1':
                header_row_index = i
                break

    if header_row_index == -1:
        print(f"⚠️ Не найдена строка заголовков в {excel_path.name}")
        return {}

    header_row = df.iloc[header_row_index]
    for i, val in enumerate(header_row):
        if pd.notna(val) and i >= 2:
            header_map[i] = str(val).strip()

    data_start_row = header_row_index + 1
    data_dict = {}

    for _, row in df.iloc[data_start_row:].iterrows():
        if row.empty or pd.isna(row.iloc[OKVED_CODE_COLUMN_INDEX]):
            continue

        okved_code = str(row.iloc[OKVED_CODE_COLUMN_INDEX]).strip()
        if not okved_code:
            continue

        # 🔍 Если фильтрация включена — пропускаем лишние
        if okved_codes_set and okved_code not in okved_codes_set:
            continue

        data_dict[okved_code] = {}

        for col_idx, key_from_header in header_map.items():
            if col_idx >= len(row):
                continue
            value = row.iloc[col_idx]
            if pd.isna(value) or str(value).strip() in ('-', '""', ''):
                formatted_value = EMPTY_CELL_MARKER
            else:
                try:
                    num_value = float(str(value).replace(',', '.').replace(' ', ''))
                    if num_value == int(num_value):
                        formatted_value = f"{int(num_value):,}".replace(',', ' ')
                    else:
                        formatted_value = f"{num_value:,.2f}".replace(',', ' ').replace('.', ',')
                except (ValueError, TypeError):
                    formatted_value = str(value).strip()
            data_dict[okved_code][key_from_header] = formatted_value

    if not data_dict:
        print(f"⚠️ В Excel-файле {excel_path.name} не найдено ни одного подходящего кода ОКВЭД.")
    else:
        print(f"📊 Загружено данных для {len(data_dict)} кодов ОКВЭД из {excel_path.name}")

    return data_dict


# === Поиск кода ОКВЭД по названию ===
def find_okved_code(cell_text, name_to_okved_cleaned):
    text_norm = _normalize_text(cell_text)
    if text_norm in name_to_okved_cleaned:
        return name_to_okved_cleaned[text_norm]
    for name, code in name_to_okved_cleaned.items():
        if text_norm in name or name in text_norm:
            print(f"⚠️ Частичное совпадение: '{cell_text}' ~ '{name}' → {code}")
            return code
    return None
