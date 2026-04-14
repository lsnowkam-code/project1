# logic.py
from pathlib import Path
import re
from config import (
    get_excel_data,
    load_okved_map,
    get_cleaned_cell_text,
    _normalize_text,
    find_okved_code
)
from docx import Document
# Исправлен импорт: добавлен WORD_COLUMNS_MAP, который будет использоваться вместо YEAR_SUFFIXES
from map_columns import WORD_COLUMNS_MAP

TAG_REGEX = re.compile(r"{{([A-Za-z0-9.-_]+)_([0-9]+)}}")


# ==========================================================
# === ГЕНЕРАЦИЯ ШАБЛОНА ====================================
# ==========================================================
def generate_word_template(input_doc_path, okved_map_path, output_doc_path):
    """
    Генерация шаблона Word с тегами {{<OKVED_CODE>_}}.
    Расширенный поиск заголовков по первым 5 строкам таблицы.
    """
    print("\n--- ШАГ 2: Генерация шаблона с умными тегами ---")
    _, name_to_okved_cleaned = load_okved_map(okved_map_path)
    doc = Document(input_doc_path)
    total_tags = 0

    for t_index, table in enumerate(doc.tables):
        print(f"\n📄 Таблица {t_index + 1}")

        # 1. Поиск заголовков в первых 5 строках
        col_to_indicator_map = {}
        for row in table.rows[:5]:
            for i, cell in enumerate(row.cells):
                header_text = _normalize_text(get_cleaned_cell_text(cell))

                # ИСПРАВЛЕНИЕ: Используем WORD_COLUMNS_MAP вместо YEAR_SUFFIXES
                for name, indicators in WORD_COLUMNS_MAP.items():
                    if _normalize_text(name) in header_text:
                        col_to_indicator_map[i] = indicators[0]  # 2022
                        if i + 1 < len(row.cells):
                            col_to_indicator_map[i + 1] = indicators[1]  # 2023
                        break

        if not col_to_indicator_map:
            print("⚠️ Заголовки не найдены. Пропускаем таблицу.")
            continue

        # 2. Вставка тегов в строки с кодами ОКВЭД
        for row_idx, row in enumerate(table.rows[2:]):  # Пропускаем заголовки
            first_cell_text = get_cleaned_cell_text(row.cells[0])
            # find_okved_code теперь импортирован в начале файла
            okved_code = find_okved_code(first_cell_text, name_to_okved_cleaned)
            if not okved_code:
                continue

            okved_tag_part = okved_code.replace('.', '_')

            for col_idx, indicator_tag_part in col_to_indicator_map.items():
                if col_idx < len(row.cells):
                    tag = f"{{{{{okved_tag_part}_{indicator_tag_part}}}}}"
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