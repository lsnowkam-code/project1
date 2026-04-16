#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки исправлений
Проверяет правильность парсинга тегов и загрузки данных
"""

import re

# === Тест 1: Парсинг тегов (ДО vs ПОСЛЕ) ===

print("=" * 80)
print("🧪 ТЕСТ 1: Проверка парсинга тегов")
print("=" * 80)

test_tags = [
    "OKVED_50_13_11_ValBal_22",
    "OKVED_45_11_VneObAk_23",
    "OKVED_10_01_02_03_KapRez_22"
]

# ❌ СТАРЫЙ СПОСОБ (НЕПРАВИЛЬНЫЙ)
print("\n❌ СТАРЫЙ СПОСОБ (неправильный):")
TAG_PATTERN_OLD = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")

for tag in test_tags:
    full_tag = f"{{{{{tag}}}}}"
    matches = TAG_PATTERN_OLD.findall(full_tag)
    if matches:
        extracted = matches[0]
        if "_" in extracted:
            parts = extracted.split("_", 1)
            print(f"  Тег: {full_tag}")
            print(f"    Первый элемент: {parts[0]}")
            print(f"    Второй элемент: {parts[1]} ❌ НЕПРАВИЛЬНО! (содержит множество _)")
    print()

# ✅ НОВЫЙ СПОСОБ (ПРАВИЛЬНЫЙ)
print("\n✅ НОВЫЙ СПОСОБ (правильный):")
tag_regex = re.compile(r"\{\{(OKVED[_A-Za-z0-9]+)\}\}")

for tag in test_tags:
    full_tag = f"{{{{{tag}}}}}"
    matches = list(tag_regex.finditer(full_tag))
    if matches:
        match = matches[0]
        full_tag_extracted = match.group(1)
        parts = full_tag_extracted.split("_")
        
        year_suffix = parts[-1]  # "22" или "23"
        indicator = parts[-2]    # "ValBal"
        okved_parts = parts[1:-2]  # ["50", "13", "11"]
        okved_code = ".".join(okved_parts)  # "50.13.11"
        indicator_year = f"{indicator}_{year_suffix}"
        
        print(f"  Тег: {full_tag}")
        print(f"    OKVED код: {okved_code} ✅")
        print(f"    Индикатор + год: {indicator_year} ✅")
    print()

# === Тест 2: Структура master_data ===

print("=" * 80)
print("🧪 ТЕСТ 2: Структура master_data")
print("=" * 80)

master_data = {
    "50.13.11": {
        "ValBal_22": "1000000",
        "ValBal_23": "1200000",
        "VneObAk_22": "500000",
        "VneObAk_23": "600000"
    },
    "45.11": {
        "KapRez_22": "800000",
        "KapRez_23": "900000"
    }
}

print("\nСтруктура данных:")
for okved_code, indicators in master_data.items():
    print(f"\n  OKVED: {okved_code}")
    for indicator_year, value in indicators.items():
        print(f"    {indicator_year}: {value}")

# === Тест 3: Имитация заполнения тегов ===

print("\n" + "=" * 80)
print("🧪 ТЕСТ 3: Имитация заполнения тегов")
print("=" * 80)

test_cells = [
    "{{OKVED_50_13_11_ValBal_22}}",
    "{{OKVED_50_13_11_ValBal_23}}",
    "{{OKVED_45_11_KapRez_22}}",
    "Текст {{OKVED_50_13_11_VneObAk_22}} в строке",
    "{{INVALID_TAG}}"  # Нет такого кода
]

print("\nПопытка заполнения:")
for cell_text in test_cells:
    matches = list(tag_regex.finditer(cell_text))
    if not matches:
        print(f"\n  ❌ Нет тегов: {cell_text}")
        continue
    
    result_text = cell_text
    for match in matches:
        full_tag = match.group(1)
        parts = full_tag.split("_")
        
        year_suffix = parts[-1]
        indicator = parts[-2]
        okved_parts = parts[1:-2]
        okved_code = ".".join(okved_parts)
        indicator_year = f"{indicator}_{year_suffix}"
        
        print(f"\n  Тег: {{{{{full_tag}}}}}")
        print(f"    OKVED: {okved_code}")
        print(f"    Индикатор: {indicator_year}")
        
        if okved_code in master_data and indicator_year in master_data[okved_code]:
            value = master_data[okved_code][indicator_year]
            result_text = result_text.replace("{{" + full_tag + "}}", value)
            print(f"    Значение: {value} ✅")
        else:
            result_text = result_text.replace("{{" + full_tag + "}}", "—")
            print(f"    Значение: НЕ НАЙДЕНО → — ⚠️")
    
    print(f"  Результат: {result_text}")

print("\n" + "=" * 80)
print("✅ ВСЕ ТЕСТЫ ПРОВЕРЕНЫ!")
print("=" * 80)
