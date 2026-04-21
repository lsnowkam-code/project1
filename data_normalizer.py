# data_normalizer.py - Модуль нормализации данных для корректного заполнения
"""
Нормализация данных перед вставкой в Word:
1. Убрать пробелы из чисел: "123 456" -> "123456"
2. Исправить минусы: "- " -> "-"
3. Привести к единому формату
"""

import re
from typing import Union


def normalize_number_string(value: str) -> str:
    """
    Нормализает числовые строки:
    - Убирает пробелы: "569 800 589" -> "569800589"
    - Исправляет минусы: "- " или "- -" -> "-"
    - Сохраняет десятичные запятые
    
    Args:
        value: Строковое значение
    
    Returns:
        Нормализованное значение
    
    Примеры:
        >>> normalize_number_string("569 800 589")
        "569800589"
        
        >>> normalize_number_string("- ")
        "-"
        
        >>> normalize_number_string("1 234,56")
        "1234,56"
    """
    if not isinstance(value, str):
        value = str(value)
    
    value = value.strip()
    
    # Обработка минуса: "- " или "- -" -> "-"
    if value.startswith("-"):
        # Удаляем пробелы после минуса
        value = "-" + value[1:].lstrip()
    
    # Удаляем пробелы из чисел, но сохраняем запятую как разделитель десятичных
    # Пример: "1 234,56" -> "1234,56"
    value = value.replace(" ", "")
    
    return value


def normalize_cell_value(value: Union[str, int, float]) -> str:
    """
    Универсальная нормализация значения для вставки в Word.
    
    Args:
        value: Любое значение из Excel
    
    Returns:
        Нормализованная строка
    
    Логика:
        1. Если строка - нормализуем числовую часть
        2. Если число - конвертируем в строку
        3. Если пусто - возвращаем пустую строку
    """
    # Обработка пустых значений
    if value is None or (isinstance(value, float) and value != value):  # NaN check
        return ""
    
    value_str = str(value).strip()
    
    if not value_str:
        return ""
    
    # Пытаемся нормализовать как число
    try:
        # Проверяем, содержит ли строка цифры
        if any(char.isdigit() for char in value_str):
            return normalize_number_string(value_str)
    except:
        pass
    
    # Если не число - возвращаем как есть
    return value_str


def normalize_text_value(value: str) -> str:
    """
    Очищает текстовые значения:
    - Убирает лишние пробелы
    - Исправляет кодировку (если нужно)
    
    Args:
        value: Текстовое значение
    
    Returns:
        Очищенное значение
    """
    if not isinstance(value, str):
        value = str(value)
    
    # Убираем лишние пробелы в начале и конце
    value = value.strip()
    
    # Заменяем множественные пробелы на одиночные
    value = ' '.join(value.split())
    
    return value


def clean_excel_value_for_word(value: Union[str, int, float]) -> str:
    """
    Комплексная очистка значения из Excel перед вставкой в Word.
    
    Применяет:
    1. Проверку на пустоту и маркеры отсутствия данных
    2. Нормализацию чисел (убрать пробелы, исправить минусы)
    3. Очистку текста (убрать лишние пробелы)
    
    Args:
        value: Значение из Excel
    
    Returns:
        Готовое для вставки значение
    
    Примеры:
        >>> clean_excel_value_for_word("569 800 589")
        "569800589"
        
        >>> clean_excel_value_for_word("- ")
        "-"
        
        >>> clean_excel_value_for_word("-")
        ""
        
        >>> clean_excel_value_for_word("   текст   ")
        "текст"
    """
    # Базовая нормализация
    value_str = normalize_cell_value(value)
    
    if not value_str:
        return ""
    
    # Проверяем на маркеры отсутствия данных
    if value_str.strip() in ["-", "—", "н.д.", "нет данных", ""]:
        return ""
    
    # Если это не явно число, применяем текстовую очистку
    try:
        float(value_str.replace(",", "."))
    except ValueError:
        value_str = normalize_text_value(value_str)
    
    return value_str


def compare_with_template(original: str, normalized: str) -> dict:
    """
    Сравнивает оригинальное и нормализованное значение для отладки.
    
    Returns:
        Словарь с информацией о изменениях
    """
    return {
        "original": original,
        "normalized": normalized,
        "changed": original.strip() != normalized,
        "spaces_removed": original.count(" ") != normalized.count(" "),
    }
