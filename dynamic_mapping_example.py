"""
Пример реализации динамического маппинга
(Это псевдокод для демонстрации концепции)
"""

import pandas as pd
from typing import Dict, Tuple, Optional
from rapidfuzz import fuzz
import unicodedata
import re

# ============================================================================
# УРОВЕНЬ 1: Идентификация источника
# ============================================================================

FILE_TYPE_PATTERNS = {
    't01': 'Баланс организаций',
    't03': 'Внеоборотные активы',
    't10': 'Оборотные активы',
    't13': 'Оборачиваемость оборотных активов',
    't19': 'Капитал и резервы',
    't20': 'Долгосрочные обязательства',
    't21': 'Краткосрочные обязательства',
}

def extract_file_type(filepath: str) -> Optional[str]:
    """Определить тип файла из названия"""
    # Паттерн: T23_000000_tXXVed14.xlsx → XX (01, 03, 10 и т.д.)
    match = re.search(r't(\d{2})Ved14', filepath)
    if match:
        return f"t{match.group(1)}"
    return None


def validate_file_type_from_row0(df: pd.DataFrame, file_type: str) -> bool:
    """Валидировать тип файла по Row[0]"""
    row0_text = ' '.join(str(cell) for cell in df.iloc[0] if pd.notna(cell))
    expected_pattern = FILE_TYPE_PATTERNS[file_type]
    
    # Проверить наличие ключевых слов и года
    has_pattern = expected_pattern.lower() in row0_text.lower()
    has_year = '2023' in row0_text or '2022' in row0_text
    
    return has_pattern and has_year


# ============================================================================
# УРОВЕНЬ 2: Локализация показателей в колонках
# ============================================================================

def parse_headers_structure(df: pd.DataFrame) -> Dict:
    """
    Парсить структуру заголовков (Row 4, 5, 6, ...)
    
    Row[4]: Основные показатели
    Row[5]: Подпоказатели (детализация)
    Row[6]: Периоды (годы)
    """
    row4 = df.iloc[4]  # Main headers
    row5 = df.iloc[5] if len(df) > 5 else None  # Sub-headers
    row6 = df.iloc[6] if len(df) > 6 else None  # Periods
    
    structure = {
        'main_headers': list(row4),
        'sub_headers': list(row5) if row5 is not None else [],
        'periods': list(row6) if row6 is not None else [],
    }
    
    return structure


def detect_year_columns(row6_data: list) -> Dict[int, str]:
    """
    Определить какие колонки содержат какие годы
    
    Returns: {col_index: year}  例: {3: '2022', 4: '2023'}
    """
    col_to_year = {}
    
    for i, cell in enumerate(row6_data):
        if pd.isna(cell):
            continue
            
        cell_str = str(cell).lower()
        
        if 'предыдущего' in cell_str or '2022' in cell_str:
            col_to_year[i] = '2022'
        elif 'отчетного' in cell_str or '2023' in cell_str:
            col_to_year[i] = '2023'
    
    return col_to_year


# ============================================================================
# УРОВЕНЬ 3: Семантическое маппинг (Текст → Код показателя)
# ============================================================================

# Словарь соответствия (можно расширять)
INDICATOR_DICTIONARY = {
    'валюта баланса': 'ValBal',
    'внеоборотные активы': 'VneObAk',
    'оборотные активы': 'ObAk',
    'оборотные активы': 'ObAkObr',
    'нематериальные активы': 'NemAk',
    'основные средства': 'OsnSr',
    'долгосрочные финансовые вложения': 'DolgFinVl',
    'краткосрочные финансовые вложения': 'FinVlKr',
    'денежные средства': 'DenSr',
    'дебиторская задолженность': 'DebZad',
    'кредиторская задолженность': 'KredZad',
    'капитал и резервы': 'KapRez',
    'уставный капитал': 'UstKap',
    'долгосрочные обязательства': 'DlOb',
    'краткосрочные обязательства': 'KrOb',
    'займы и кредиты': 'ZaimKrDl',
}


def normalize_text(text: str) -> str:
    """
    Нормализовать текст для сравнения
    - Нижний регистр
    - Удалить диакритические знаки (ё → е)
    - Удалить скобки и пояснения
    """
    if not isinstance(text, str) or pd.isna(text):
        return ""
    
    # Нижний регистр
    text = text.lower()
    
    # Удалить диакритические знаки
    text = unicodedata.normalize('NFD', text)
    text = ''.join(char for char in text if unicodedata.category(char) != 'Mn')
    
    # Удалить скобки и всё что в них
    text = re.sub(r'\([^)]*\)', '', text)
    
    # Сокращать кратные пробелы
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text


def find_indicator_code(text: str, confidence_threshold: float = 0.75) -> Tuple[Optional[str], float]:
    """
    Найти код показателя по тексту
    
    Returns: (indicator_code, confidence_score)
    """
    normalized = normalize_text(text)
    
    if not normalized:
        return None, 0.0
    
    # Первый проход: точное совпадение
    if normalized in INDICATOR_DICTIONARY:
        return INDICATOR_DICTIONARY[normalized], 1.0
    
    # Второй проход: частичное совпадение (longest match)
    best_match = None
    best_score = 0
    
    for dict_key, indicator_code in INDICATOR_DICTIONARY.items():
        # Найти самое длинное совпадение (если ключ содержится в тексте)
        if dict_key in normalized:
            score = len(dict_key) / len(normalized)  # Доля совпадения
            if score > best_score:
                best_score = score
                best_match = indicator_code
    
    if best_score > 0.7:
        return best_match, best_score
    
    # Третий проход: fuzzy matching (Левенштейна)
    for dict_key, indicator_code in INDICATOR_DICTIONARY.items():
        score = fuzz.ratio(normalized, dict_key) / 100.0
        
        if score > best_score and score >= confidence_threshold:
            best_score = score
            best_match = indicator_code
    
    return best_match, best_score


# ============================================================================
# УРОВЕНЬ 4: Построение маппинга
# ============================================================================

def extract_mapping_from_excel(filepath: str, excel_file_key: str) -> Dict:
    """
    Полный алгоритм экстракции маппинга из Excel
    
    Returns:
    {
        ("T23_000000_t03", "долгосрочные финансовые вложения"): {
            "indicator": "DolgFinVl",
            "year_2022": 17,
            "year_2023": 18,
            "confidence": 0.95,
            "validated": True
        }
    }
    """
    
    # Шаг 1: Определить тип файла
    file_type = extract_file_type(filepath)
    if not file_type:
        print(f"❌ Не удалось определить тип файла: {filepath}")
        return {}
    
    print(f"✓ Тип файла: {file_type}")
    
    # Шаг 2: Прочитать Excel
    try:
        df = pd.read_excel(filepath, header=None)
    except Exception as e:
        print(f"❌ Ошибка чтения Excel: {e}")
        return {}
    
    # Шаг 3: Валидировать по Row[0]
    if not validate_file_type_from_row0(df, file_type):
        print(f"⚠️  Row[0] не подтверждает тип {file_type}")
    
    # Шаг 4: Парсить структуру заголовков
    structure = parse_headers_structure(df)
    print(f"✓ Найдено основных показателей: {len(structure['main_headers'])}")
    
    # Шаг 5: Определить колонки периодов (года)
    col_to_year = detect_year_columns(structure['periods'])
    print(f"✓ Колонки годов: {col_to_year}")
    
    # Шаг 6: Построить маппинг
    mapping = {}
    
    # Проходим по каждой колонке
    for col_idx, year in col_to_year.items():
        if col_idx >= len(structure['main_headers']):
            continue
        
        main_header = structure['main_headers'][col_idx]
        sub_header = structure['sub_headers'][col_idx] if col_idx < len(structure['sub_headers']) else None
        
        # Приоритет: sub_header (детальный показатель), потом main_header
        text_to_match = str(sub_header) if sub_header and str(sub_header) != 'nan' else str(main_header)
        
        # Найти код показателя
        indicator_code, confidence = find_indicator_code(text_to_match)
        
        if indicator_code:
            # Нормализованный ключ для маппинга
            normalized_text = normalize_text(text_to_match)
            key = (excel_file_key, normalized_text)
            
            # Если это первое появление показателя (год 2022)
            if year == '2022':
                mapping[key] = {
                    "indicator": indicator_code,
                    "year_2022": col_idx,
                    "year_2023": None,  # Заполнится позже
                    "confidence": confidence,
                    "main_header": str(main_header),
                    "sub_header": str(sub_header) if sub_header else None,
                }
            # Если уже есть запись (добавляем 2023)
            elif key in mapping and year == '2023':
                mapping[key]["year_2023"] = col_idx
    
    # Шаг 7: Валидировать данные (Row[8+])
    if len(df) > 8:
        sample_row = df.iloc[8]  # Первая строка данных
        for key in mapping:
            col_2022 = mapping[key]["year_2022"]
            if col_2022 < len(sample_row):
                value = sample_row[col_2022]
                if pd.notna(value) and str(value).strip():
                    mapping[key]["validated"] = True
                    mapping[key]["confidence"] = min(1.0, mapping[key]["confidence"] + 0.05)
    
    return mapping


# ============================================================================
# Пример использования
# ============================================================================

if __name__ == "__main__":
    # Пример: обработать файл T23_000000_t03Ved14.xlsx
    filepath = "input/excel/T23_000000_t03Ved14.xlsx"
    
    print("=" * 70)
    print("ДИНАМИЧЕСКОЕ ИЗВЛЕЧЕНИЕ МАППИНГА ИЗ EXCEL")
    print("=" * 70)
    
    mapping = extract_mapping_from_excel(filepath, "T23_000000_t03Ved14")
    
    print("\n📊 РЕЗУЛЬТАТЫ МАППИНГА:")
    print("-" * 70)
    
    for (file_key, indicator_text), details in mapping.items():
        print(f"\n🔹 {indicator_text}")
        print(f"   Показатель: {details['indicator']}")
        print(f"   Колонка 2022: {details['year_2022']}")
        print(f"   Колонка 2023: {details['year_2023']}")
        print(f"   Уверенность: {details['confidence']:.2%}")
        print(f"   Валидировано: {details.get('validated', False)}")

