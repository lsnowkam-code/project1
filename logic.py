# logic.py
from pathlib import Path
import re
from config import (
    get_excel_data,
    load_okved_map,
    load_table_source_map,
    load_column_mapping,
    get_cleaned_cell_text,
    _normalize_text,
    find_okved_code,
    get_table_name
)
from config_v2 import load_column_mapping_v2
from docx import Document

TAG_REGEX = re.compile(r"{{([^}]+?)_([0-9]+)}}")


# ==========================================================
# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==============================
# ==========================================================
def auto_detect_table_source(table, table_source_mapping, word_to_indicator):
    """
    Автоматически определяет источник Excel файла для таблицы 
    по её содержимому (проверяет наличие ключевых показателей).
    """
    # Собираем текст из первых 5 строк таблицы
    table_content = ' '.join([
        get_cleaned_cell_text(cell).lower()
        for row in table.rows[:5]
        for cell in row.cells
    ])
    
    # Пробуем найти текстовые совпадения показателей в таблице
    for indicator_name in word_to_indicator.keys():
        indicator_name_lower = indicator_name.lower()
        if indicator_name_lower in table_content:
            # Нашли совпадение! Нужно найти файл для этого показателя
            # word_to_indicator[indicator_name] = (file, indicator_code)
            file_info = word_to_indicator.get(indicator_name)
            if isinstance(file_info, tuple) and len(file_info) >= 1:
                return file_info[0]  # Возвращаем файл
            elif isinstance(file_info, str):
                return file_info  # Если просто строка
    
    return None


# ==========================================================
# === ГЕНЕРАЦИЯ ШАБЛОНА ====================================
# ==========================================================
def generate_word_template(input_doc_path, okved_map_path, table_source_mapping_path, column_mapping_path, output_doc_path):
    """
    Генерация шаблона Word с тегами {{OKVED_<code>_<indicator>[_22|_23]}}.
    Расширенный поиск заголовков по первым 5 строкам таблицы.
    """
    print("\n--- ШАГ 2: Генерация шаблона с умными тегами ---")
    _, name_to_okved_cleaned = load_okved_map(okved_map_path)
    table_source_mapping = load_table_source_map(table_source_mapping_path)
    word_to_indicator, _, indicator_to_file, file_word_to_indicator, _ = load_column_mapping_v2(column_mapping_path)
    doc = Document(input_doc_path)
    total_tags = 0
    current_source_file = None
    normalized_title_to_src = {k: v for k, v in table_source_mapping.items()}

    for t_index, table in enumerate(doc.tables):
        print(f"\n📄 Таблица {t_index + 1}")

        # СПЕЦИАЛЬНЫЙ СЛУЧАЙ: Таблица 17 - это таблица 7 (Долгосрочные обязательства)
        if t_index + 1 == 17:
            print(f"⚠️ Таблица 17 - применяем специальное определение (таблица 7)")
            current_source_file = 'T23_000000_t20Ved14.xlsx'
            print(f"   → Установлен источник: {current_source_file}")
        else:
            # 0. Определение источника данных по названию таблицы
            table_title = get_table_name(table, table_source_mapping.keys())
            if table_title:
                table_title_norm = _normalize_text(table_title)
                if table_title_norm in normalized_title_to_src:
                    current_source_file = normalized_title_to_src[table_title_norm]
                    print(f"🔍 Источник таблицы: {current_source_file}")
                else:
                    print(f"⚠️ Не найден источник для заголовка таблицы: '{table_title}'")
                    # Fallback: try to auto-detect by table content
                    print(f"   → Пытаемся определить по содержимому таблицы...")
                    detected = auto_detect_table_source(table, table_source_mapping, word_to_indicator)
                    if detected:
                        current_source_file = detected
                        print(f"   ✓ Автоматически определен источник: {current_source_file}")
            else:
                print("ℹ️ Заголовок таблицы не определён, используем предыдущий источник")

        if not current_source_file:
            print("⚠️ Источник не определён. Пропускаем таблицу.")
            continue

        # Используем file_word_to_indicator для правильного маппирования по (файл, слово)
        source_word_to_indicator = {
            name: indicator
            for (file, name), indicator in file_word_to_indicator.items()
            if file == current_source_file
        }

        if not source_word_to_indicator:
            print(f"⚠️ Нет показателей для источника {current_source_file}. Пропускаем таблицу.")
            continue

        # DEBUG для таблицы 17
        if t_index + 1 == 17:
            print(f"   DEBUG: Таблица 17")
            print(f"   Источник: {current_source_file}")
            print(f"   Показатели ({len(source_word_to_indicator)}):")
            for name, indicator in list(source_word_to_indicator.items())[:5]:
                print(f"      {name} → {indicator}")

        # 1. Находим строку с годами (2022 / 2023) в зоне заголовка
        year_row = None
        year_row_idx = None
        header_scan_limit = min(8, len(table.rows))
        for row_idx, row in enumerate(table.rows[:header_scan_limit]):
            row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
            if any(year in ('2022', '2023') for year in row_years):
                year_row = row
                year_row_idx = row_idx
                break

        # 2. Строим маппинг: столбец -> (показатель, год)
        # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ:
        # Используем несколько строк заголовка до year_row, а не одну indicator_row.
        col_to_indicator_map = {}
        if year_row is not None and year_row_idx is not None:
            # Собираем "сетку" заголовков с horizontal forward-fill для объединённых ячеек.
            header_rows_filled = []
            for row in table.rows[:year_row_idx]:
                row_values = []
                last_non_empty = ""
                for cell in row.cells:
                    value = _normalize_text(get_cleaned_cell_text(cell))
                    if value:
                        last_non_empty = value
                    else:
                        value = last_non_empty
                    row_values.append(value)
                header_rows_filled.append(row_values)

            # Для каждого столбца года пытаемся найти лучший indicator по составному заголовку.
            last_indicator = None
            for i, cell in enumerate(year_row.cells):
                year_text = get_cleaned_cell_text(cell).strip()
                if year_text not in ('2022', '2023'):
                    continue

                parts = []
                seen = set()
                for header_row in header_rows_filled:
                    if i >= len(header_row):
                        continue
                    header_part = header_row[i].strip()
                    if not header_part or header_part in seen:
                        continue
                    seen.add(header_part)
                    parts.append(header_part)
                composed_header = " ".join(parts)

                best_match = None
                best_len = 0
                for name, indicator in source_word_to_indicator.items():
                    if name in composed_header and len(name) > best_len:
                        best_match = indicator
                        best_len = len(name)

                # Fallback: если в конкретном столбце не нашли показатель,
                # используем последний найденный (типичный случай пары 2022/2023).
                if best_match:
                    last_indicator = best_match
                elif last_indicator:
                    best_match = last_indicator

                if not best_match:
                    continue

                col_to_indicator_map[i] = (best_match, '22' if year_text == '2022' else '23')
                print(f"   🔍 Столбец {i}: '{composed_header[:90]}' -> {best_match}, год {year_text}")

        if not col_to_indicator_map:
            # Fallback: старая логика для нестандартных таблиц.
            base_indicator_map = {}
            last_indicator = None
            for row in table.rows[:5]:
                for i, cell in enumerate(row.cells):
                    header_text = _normalize_text(get_cleaned_cell_text(cell))
                    for name, indicator in source_word_to_indicator.items():
                        if name in header_text:
                            base_indicator_map[i] = indicator
                            last_indicator = indicator
                            break
                    else:
                        if last_indicator is not None and not header_text:
                            base_indicator_map[i] = last_indicator

            if year_row:
                for i, cell in enumerate(year_row.cells):
                    year_text = get_cleaned_cell_text(cell).strip()
                    indicator = base_indicator_map.get(i)
                    if indicator is None:
                        continue
                    if year_text == '2022':
                        col_to_indicator_map[i] = (indicator, '22')
                    elif year_text == '2023':
                        col_to_indicator_map[i] = (indicator, '23')
            else:
                for i, indicator in base_indicator_map.items():
                    col_to_indicator_map[i] = (indicator, None)

        if not col_to_indicator_map:
            print("⚠️ Заголовки не найдены. Пропускаем таблицу.")
            continue

        # 4. Вставка тегов в строки с кодами ОКВЭД
        for row_idx, row in enumerate(table.rows[2:]):  # Пропускаем заголовки
            first_cell_text = get_cleaned_cell_text(row.cells[0])
            okved_code = find_okved_code(first_cell_text, name_to_okved_cleaned)
            if not okved_code:
                continue

            okved_tag_part = okved_code.replace('.', '_')

            for col_idx, indicator_spec in col_to_indicator_map.items():
                if col_idx < len(row.cells):
                    indicator, year = indicator_spec
                    if year:
                        tag = f"{{{{OKVED_{okved_tag_part}_{indicator}_{year}}}}}"
                    else:
                        tag = f"{{{{OKVED_{okved_tag_part}_{indicator}}}}}"
                    # Очищаем ячейку перед вставкой тега
                    for p in row.cells[col_idx].paragraphs:
                        p.text = " "
                    if row.cells[col_idx].paragraphs:
                        row.cells[col_idx].paragraphs[0].text = tag
                    total_tags += 1
                    print(f"   🏷️ Строка {row_idx + 2}: вставлен тег {tag}")

    doc.save(output_doc_path)
    print(f"\n✅ Шаблон с тегами сохранён: {output_doc_path}")
    print(f"🔢 Всего вставлено тегов: {total_tags}")


# ==========================================================
# === ДИАГНОСТИКА И ПОМОЩНЫЕ ==============================
# ==========================================================
def find_unfilled_tags(doc_path):
    """Диагностика незаполненных тегов."""
    doc = Document(doc_path)
    unfilled_tags = set()
    for para in doc.paragraphs:
        matches = TAG_REGEX.findall(para.text)
        for code, index in matches:
            unfilled_tags.add(f"{{{{{code}_{index}}}}}")

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    matches = TAG_REGEX.findall(para.text)
                    for code, index in matches:
                        unfilled_tags.add(f"{{{{{code}_{index}}}}}")

    print(f"\n🔍 Незаполненные теги: {len(unfilled_tags)}")
    for tag in sorted(unfilled_tags):
        print(f" - {tag}")
    return unfilled_tags


def collect_okved_codes_from_template(doc_path):
    """Собирает коды ОКВЭД из тегов шаблона."""
    doc = Document(doc_path)
    codes = set()
    for para in doc.paragraphs:
        matches = TAG_REGEX.findall(para.text)
        for code, _ in matches:
            codes.add(code)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    matches = TAG_REGEX.findall(para.text)
                    for code, _ in matches:
                        codes.add(code)
    print(f"\n📄 Коды ОКВЭД в шаблоне: {len(codes)}")
    return codes


def collect_okved_codes_from_excel(excel_dir, table_mapping, okved_codes_set):
    """Собирает коды ОКВЭД, реально присутствующие в Excel."""
    all_codes = set()
    for _, excel_filename in table_mapping.items():
        excel_path = Path(excel_dir) / excel_filename
        if not excel_path.exists():
            print(f"⚠️ Excel-файл не найден: {excel_path}")
            continue
        try:
            data = get_excel_data(excel_path, okved_codes_set)
            all_codes.update(data.keys())
        except Exception as e:
            print(f"⚠️ Ошибка при чтении {excel_filename}: {e}")
    print(f"\n📊 Коды ОКВЭД в Excel: {len(all_codes)}")
    return all_codes


def compare_okved_sets(template_codes, excel_codes):
    """Сравнение множеств кодов ОКВЭД."""
    missing = template_codes - excel_codes
    extra = excel_codes - template_codes
    print("\n🔍 В шаблоне, но нет в Excel:")
    for code in sorted(missing):
        print(f" - {code}")
    print("\n📁 В Excel, но не используется в шаблоне:")
    for code in sorted(extra):
        print(f" - {code}")
