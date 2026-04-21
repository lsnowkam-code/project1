# confic.py
import csv
import re

import pandas as pd
from docx.table import _Cell, Table

# === Константы ===
EMPTY_CELL_MARKER = "-"  # Обычный дефис вместо длинного тире
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
    - file_word_to_indicator: (файл, нормализованное название) → код индикатора
    
    ИСПРАВЛЕНО: Исправлен дубликат кода SobAkc для Нераспределенной прибыли
    """
    word_to_indicator = {}
    indicator_to_excel = {}
    indicator_to_file = {}
    file_word_to_indicator = {}

    with open(filepath, encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=';')
        headers = next(reader)  # Пропускаем заголовок

        for row in reader:
            if len(row) < 5:
                continue  # Пропускаем неполные строки

            excel_file = str(row[0]).strip()
            word_name = str(row[1]).strip()
            indicator = str(row[2]).strip()
            col_22 = str(row[3]).strip() if len(row) > 3 else ''
            col_23 = str(row[4]).strip() if len(row) > 4 else ''

            # 🔧 ИСПРАВЛЕНИЕ ДУБЛИКАТА: Собственные акции vs Нераспределенная прибыль
            if "Нераспределенная" in word_name and indicator == "SobAkc":
                print(f"⚠️ АВТО-ИСПРАВЛЕНИЕ: Код для '{word_name}' изменён с SobAkc на NeraspPribyl")
                indicator = "NeraspPribyl"
            
            if col_22.lower() == 'nan':
                col_22 = ''
            if col_23.lower() == 'nan':
                col_23 = ''

            word_to_indicator[word_name] = indicator
            indicator_to_excel[indicator] = (col_22, col_23)
            indicator_to_file[indicator] = excel_file
            # NEW: Added mapping per file to avoid conflicts when same indicator in multiple files
            file_word_to_indicator[(excel_file, _normalize_text(word_name))] = indicator

    return word_to_indicator, indicator_to_excel, indicator_to_file, file_word_to_indicator


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
def get_excel_data(excel_path, okved_codes_set, file_word_to_indicator=None):
    try:
        df = pd.read_excel(excel_path, header=None, dtype=str)
    except Exception as e:
        print(f"❌ Ошибка чтения Excel: {e}")
        return {}

    header_row_index = -1
    
    # 🔍 ДИНАМИЧЕСКИЙ ПОИСК заголовка: ищем строку с названиями показателей (содержит "Код" или "Наименование")
    for i, row in df.head(20).iterrows():
        if len(row) > 1 and pd.notna(row.iloc[0]):
            cell_text = str(row.iloc[0]).strip().lower()
            if 'код' in cell_text or 'наименование' in cell_text:
                header_row_index = i
                break
    
    # Fallback: старый метод поиска по 'А' и '1'
    if header_row_index == -1:
        for i, row in df.head(20).iterrows():
            if len(row) > 2 and pd.notna(row.iloc[0]) and pd.notna(row.iloc[2]):
                if str(row.iloc[0]).strip().upper() == 'А' and str(row.iloc[2]).strip() == '1':
                    header_row_index = i
                    break

    if header_row_index == -1:
        print(f"⚠️ Не найдена строка заголовков в {excel_path.name}")
        return {}

    print(f"📍 Заголовок найден в строке {header_row_index + 1}")
    
    # 🔍 Поиск строки с периодами (ищем ДО или ПОСЛЕ строки заголовка)
    periods_row_index = -1
    
    # Сначала ищем ПОСЛЕ заголовка (до 10 строк вперед)
    for offset in range(1, 11):
        check_idx = header_row_index + offset
        if check_idx >= len(df):
            break
        row = df.iloc[check_idx]
        for val in row[:20]:  # Проверяем больше колонок
            if pd.notna(val) and ('предыдущего' in str(val) or 'отчетного' in str(val)):
                periods_row_index = check_idx
                break
        if periods_row_index != -1:
            break
    
    # Если не нашли после, ищем ДО заголовка
    if periods_row_index == -1:
        for offset in range(1, 11):
            check_idx = header_row_index - offset
            if check_idx < 0:
                break
            row = df.iloc[check_idx]
            for val in row[:20]:
                if pd.notna(val) and ('предыдущего' in str(val) or 'отчетного' in str(val)):
                    periods_row_index = check_idx
                    break
            if periods_row_index != -1:
                break
    
    if periods_row_index != -1:
        print(f"📍 Строка периодов найдена в строке {periods_row_index + 1}")
    
    # Построение маппинга колонок: номер колонки → (показатель, год)
    column_mapping = {}  # col_idx -> {'indicator': name, 'year': '22'/'23'}
    
    # Определяем названия показателей из строки заголовка
    header_row = df.iloc[header_row_index]
    indicators_map = {}  # col_idx -> indicator_name
    
    # Проверяем, есть ли дополнительные строки с подзаголовками (как в t19Ved14.xlsx строка 5)
    subheader_row_index = -1
    if header_row_index + 1 < len(df):
        check_row = df.iloc[header_row_index + 1]
        # Если в этой строке есть текст, но нет периодов - это подзаголовки
        has_text = any(pd.notna(check_row.iloc[i]) and 'предыдущего' not in str(check_row.iloc[i]).lower() and 'отчетного' not in str(check_row.iloc[i]).lower() for i in range(2, min(20, len(check_row))))
        has_periods = any(pd.notna(check_row.iloc[i]) and ('предыдущего' in str(check_row.iloc[i]).lower() or 'отчетного' in str(check_row.iloc[i]).lower()) for i in range(2, min(20, len(check_row))))
        if has_text and not has_periods:
            subheader_row_index = header_row_index + 1
            print(f"📍 Найдена строка подзаголовков в строке {subheader_row_index + 1}")
    
    # Сначала заполняем явные названия показателей из основной строки заголовка
    last_indicator = None
    for i, val in enumerate(header_row):
        if pd.notna(val) and i >= 2:  # Пропускаем первые 2 колонки (Код, Наименование)
            text = str(val).strip()
            if text and 'предыдущего' not in text and 'отчетного' not in text and 'итог' not in text.lower():
                indicators_map[i] = text
                last_indicator = text
    
    # Если есть подзаголовки, используем их для колонок без явного заголовка
    if subheader_row_index != -1:
        subheader_row = df.iloc[subheader_row_index]
        for i, val in enumerate(subheader_row):
            if pd.notna(val) and i >= 2 and i not in indicators_map:
                text = str(val).strip()
                if text and 'предыдущего' not in text and 'отчетного' not in text:
                    indicators_map[i] = text
                    # Обновляем last_indicator для последующих колонок
                    last_indicator = text
    
    # Заполняем оставшиеся колонки последним известным показателем
    # Это нужно для случаев когда показатели идут через колонку (как в t01Ved14.xlsx)
    current_indicator = None
    for i in range(2, len(header_row)):
        if i in indicators_map:
            current_indicator = indicators_map[i]
        elif current_indicator is not None:
            indicators_map[i] = current_indicator
    
    # Определяем периоды из строки периодов
    # ВАЖНО: Обрабатываем ВСЕ колонки с периодами, даже если название показателя в соседней колонке
    if periods_row_index != -1:
        periods_row = df.iloc[periods_row_index]
        
        # Сначала проходим по всем колонкам и определяем периоды
        current_indicator = None
        for i, val in enumerate(periods_row):
            if pd.notna(val) and i >= 2:
                text = str(val).strip().lower()
                year = None
                if 'предыдущего' in text:
                    year = '22'
                elif 'отчетного' in text:
                    year = '23'
                
                # Если нашли период, пытаемся определить показатель
                if year:
                    # Сначала проверяем, есть ли явное название в этой колонке
                    indicator_name = indicators_map.get(i)
                    
                    # Если нет, используем последний известный показатель (для случаев like t01Ved14.xlsx)
                    if not indicator_name and current_indicator:
                        indicator_name = current_indicator
                    
                    # Если всё ещё нет, пробуем найти в подзаголовках
                    if not indicator_name and subheader_row_index != -1:
                        sub_val = df.iloc[subheader_row_index, i]
                        if pd.notna(sub_val):
                            sub_text = str(sub_val).strip()
                            if sub_text and 'предыдущего' not in sub_text.lower() and 'отчетного' not in sub_text.lower():
                                indicator_name = sub_text
                    
                    if indicator_name:
                        column_mapping[i] = {
                            'indicator': _normalize_text(indicator_name),
                            'year': year,
                            'original_name': indicator_name
                        }
                        current_indicator = indicator_name
    
    data_start_row = max(header_row_index, periods_row_index) + 1 if periods_row_index != -1 else header_row_index + 1
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

        for col_idx, mapping_info in column_mapping.items():
            if col_idx >= len(row):
                continue
            
            value = row.iloc[col_idx]
            if pd.isna(value) or str(value).strip() in ('-', '""', ''):
                formatted_value = EMPTY_CELL_MARKER
            else:
                try:
                    # Удаляем пробелы из числа перед конвертацией (569 800 589 → 569800589)
                    num_value = float(str(value).replace(' ', '').replace(',', '.').replace(' ', ''))
                    if num_value == int(num_value):
                        # Форматируем БЕЗ пробелов-разделителей тысяч
                        formatted_value = f"{int(num_value)}"
                    else:
                        formatted_value = f"{num_value:.2f}".replace('.', ',')
                except (ValueError, TypeError):
                    formatted_value = str(value).strip()
            
            # 🔧 КОНВЕРТАЦИЯ: нормализованное название → код индикатора
            indicator_code = None
            if file_word_to_indicator:
                key = (excel_path.name, mapping_info['indicator'])
                indicator_code = file_word_to_indicator.get(key)
            
            # Если не нашли код или маппинг не передан, используем нормализованное название (fallback)
            if indicator_code is None:
                indicator_code = mapping_info['indicator']
            
            # Ключ формата: INDICATOR_YY (например, ValBal_22)
            indicator_key = f"{indicator_code}_{mapping_info['year']}"
            data_dict[okved_code][indicator_key] = formatted_value

    if not data_dict:
        print(f"⚠️ В Excel-файле {excel_path.name} не найдено ни одного подходящего кода ОКВЭД.")
    else:
        print(f"📊 Загружено данных для {len(data_dict)} кодов ОКВЭД из {excel_path.name}")
        if column_mapping:
            print(f"   Найдено колонок: {len(column_mapping)}")
            # Отладочная информация
            for col_idx, info in list(column_mapping.items())[:5]:
                indicator_code = file_word_to_indicator.get((excel_path.name, info['indicator']), info['indicator']) if file_word_to_indicator else info['indicator']
                print(f"   Колонка {col_idx}: {info['original_name']} → {indicator_code}_{info['year']}")

    return data_dict


# === Поиск кода ОКВЭД по названию ===
def find_okved_code(cell_text, name_to_okved_cleaned):
    text_norm = _normalize_text(cell_text)
    if not text_norm or text_norm in ('код', 'наименование'):
        return None
    if text_norm in name_to_okved_cleaned:
        return name_to_okved_cleaned[text_norm]
    for name, code in name_to_okved_cleaned.items():
        if len(text_norm) >= 3 and (text_norm in name or name in text_norm):
            print(f"⚠️ Частичное совпадение: '{cell_text}' ~ '{name}' → {code}")
            return code
    return None
