# test_smart_loader.py - Тестирование функций умного поиска
"""
Тесты для демонстрации работы:
1. Нечеткого совпадения (fuzzy_match)
2. Динамического поиска колонок по году
3. Нечеткого поиска строк
4. Нормализации данных
"""

import pandas as pd
from smart_loader import (
    normalize_text,
    fuzzy_match,
    find_column_by_year,
    find_row_by_fuzzy_match
)
from data_normalizer import (
    normalize_number_string,
    clean_excel_value_for_word
)


def test_normalize_text():
    """Тест: Нормализация текста"""
    print("\n🧪 Тест 1: Нормализация текста")
    print("=" * 50)
    
    test_cases = [
        ("  Внеоборотные   активы  ", "внеоборотные активы"),
        ("КАПИТАЛ И РЕЗЕРВЫ", "капитал и резервы"),
        ("на конец", "на конец"),
    ]
    
    for input_text, expected in test_cases:
        result = normalize_text(input_text)
        status = "✅" if result == expected else "❌"
        print(f"{status} '{input_text}' → '{result}' (ожидалось: '{expected}')")


def test_fuzzy_match():
    """Тест: Нечеткое совпадение"""
    print("\n🧪 Тест 2: Нечеткое совпадение (fuzzy_match)")
    print("=" * 50)
    
    candidates = [
        "Внеоборотные активы (тыс.руб)",
        "Оборотные активы",
        "Капитал и резервы"
    ]
    
    test_cases = [
        ("Внеоборотные активы", 0.80),  # Должно найтись с высоким сходством
        ("активы", 0.80),  # Может частично совпадать
        ("капитал", 0.80),  # Капитал и резервы
    ]
    
    for target, threshold in test_cases:
        result, score = fuzzy_match(target, candidates, threshold=threshold)
        if result:
            print(f"✅ '{target}' найден как '{result}' (сходство: {score:.1%})")
        else:
            print(f"❌ '{target}' не найден")


def test_find_column_by_year():
    """Тест: Поиск колонки по году"""
    print("\n🧪 Тест 3: Динамический поиск колонки по году")
    print("=" * 50)
    
    # Создаем тестовый DataFrame
    df = pd.DataFrame({
        0: ["Показатель", "Внеоборотные активы", "Оборотные активы"],
        1: ["2022", "100", "200"],
        2: ["2023", "150", "250"],
        3: ["На начало", "X", "Y"],
    })
    
    df_names = {0: "Показатель", 1: "2022", 2: "2023", 3: "На начало"}
    
    test_cases = [
        ("2023", []),
        ("2022", []),
    ]
    
    for year, markers in test_cases:
        col_idx, header = find_column_by_year(df, year, year_markers=markers)
        if col_idx is not None:
            print(f"✅ Год '{year}' найден в колонке {col_idx}: '{header}'")
        else:
            print(f"❌ Год '{year}' не найден")


def test_find_row_by_fuzzy_match():
    """Тест: Нечеткий поиск строк"""
    print("\n🧪 Тест 4: Нечеткий поиск строк по совпадению")
    print("=" * 50)
    
    # Создаем тестовый DataFrame
    df = pd.DataFrame({
        0: ["Показатель", "Внеоборотные активы (тыс.руб)", "Оборотные активы", "Капитал и резервы"],
        1: [100, 200, 300, 400],
    })
    
    test_cases = [
        ("Внеоборотные активы", 0.80),
        ("активы", 0.70),
    ]
    
    for target, threshold in test_cases:
        row_idx, row_name, score = find_row_by_fuzzy_match(df, target, threshold=threshold)
        if row_idx is not None:
            print(f"✅ '{target}' найден в строке {row_idx}: '{row_name}' (сходство: {score:.1%})")
        else:
            print(f"❌ '{target}' не найден")


def test_normalize_numbers():
    """Тест: Нормализация чисел"""
    print("\n🧪 Тест 5: Нормализация числовых данных")
    print("=" * 50)
    
    test_cases = [
        ("569 800 589", "569800589"),
        ("1 234,56", "1234,56"),
        ("- ", "-"),
        ("-", "-"),
        ("123", "123"),
    ]
    
    for input_val, expected in test_cases:
        result = normalize_number_string(input_val)
        status = "✅" if result == expected else "❌"
        print(f"{status} '{input_val}' → '{result}' (ожидалось: '{expected}')")


def test_clean_excel_value():
    """Тест: Комплексная очистка значений"""
    print("\n🧪 Тест 6: Комплексная очистка значений из Excel")
    print("=" * 50)
    
    test_cases = [
        ("569 800 589", "569800589"),
        ("  текст  ", "текст"),
        ("- ", "-"),
        (None, ""),
        (123.45, "123.45"),
    ]
    
    for input_val, expected in test_cases:
        result = clean_excel_value_for_word(input_val)
        status = "✅" if result == expected else "❌"
        print(f"{status} {repr(input_val)} → '{result}' (ожидалось: '{expected}')")


def run_all_tests():
    """Запуск всех тестов"""
    print("\n" + "=" * 50)
    print("🧪 ЗАПУСК ТЕСТОВ УМНОГО ПОИСКА")
    print("=" * 50)
    
    test_normalize_text()
    test_fuzzy_match()
    test_find_column_by_year()
    test_find_row_by_fuzzy_match()
    test_normalize_numbers()
    test_clean_excel_value()
    
    print("\n" + "=" * 50)
    print("✅ ТЕСТЫ ЗАВЕРШЕНЫ")
    print("=" * 50)


if __name__ == "__main__":
    run_all_tests()
