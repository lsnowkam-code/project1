#!/usr/bin/env python3
"""Анализирует все таблицы в эталонном документе и определяет их названия"""

from docx import Document
import re

doc = Document('output/как должно быть!.docx')

# Ищверём названия таблиц
def get_table_title(table):
    """Получает название таблицы из первой строки или найденного текста рядом"""
    # Первая строка
    if table.rows:
        first_row_text = ' '.join([cell.text.strip() for cell in table.rows[0].cells]).strip()
        if first_row_text and len(first_row_text) > 10:
            return first_row_text
    
    # Вторая строка
    if len(table.rows) > 1:
        second_row_text = ' '.join([cell.text.strip() for cell in table.rows[1].cells]).strip()
        if second_row_text and len(second_row_text) > 10:
            return second_row_text
    
    return None

# Вывод всех таблиц
print("АНАЛИЗ ВСЕХ ТАБЛИЦ В ЭТАЛОННОМ ДОКУМЕНТЕ")
print("=" * 80)

for t_idx, table in enumerate(doc.tables):
    title = get_table_title(table)
    rows = len(table.rows)
    cols = len(table.rows[0].cells) if table.rows else 0
    
    print(f"\nТаблица {t_idx + 1}:")
    print(f"  Размер: {rows}x{cols}")
    print(f"  Название: {title}")
    if table.rows and len(table.rows) > 1:
        second_row = ' '.join([c.text[:30] for c in table.rows[1].cells[:3]]).strip()
        print(f"  Вторая строка (первые 3 ячейки): {second_row}")
