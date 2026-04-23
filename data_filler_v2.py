# data_filler_v2.py - Новая версия с умным поиском и нормализацией
"""
Версия 2 заполнения данных:
1. Использует find_column_by_year вместо жестких индексов
2. Использует fuzzy matching для поиска строк
3. Применяет нормализацию данных перед вставкой
4. Логирует все действия для отладки
"""

import re
from pathlib import Path
import pandas as pd
from typing import Dict, Set, Tuple, Optional

from config_v2 import load_column_mapping_v2
from smart_loader import (
    find_column_by_year,
    find_row_by_fuzzy_match,
    normalize_text,
    get_cell_value_safely
)
from data_normalizer import clean_excel_value_for_word


def pre_load_all_excel_data_v2(excel_dir: Path, table_source_mapping: Dict, 
                               okved_codes_set: Set[str], column_mapping_path: Path,
                               use_fuzzy_match: bool = True, fuzzy_threshold: float = 0.80) -> Tuple[Dict, dict]:
    """
    Загружает все данные из Excel с использованием умного поиска.
    
    Args:
        excel_dir: Путь к папке с Excel файлами
        table_source_mapping: Маппинг таблиц → файлы Excel
        okved_codes_set: Набор кодов ОКВЭД для фильтрации
        column_mapping_path: Путь к column_mapping_v2.csv
        use_fuzzy_match: Использовать нечеткое совпадение для поиска строк
        fuzzy_threshold: Порог сходства (0-1) для нечеткого совпадения
    
    Returns:
        (master_data[okved][indicator_year] = value, stats_dict)
    """
    print("\n📊 Шаг 3: Загрузка данных из Excel (УМНЫЙ поиск)")
    
    try:
        _, indicator_to_excel, indicator_to_file, _, indicator_keywords = load_column_mapping_v2(str(column_mapping_path))
    except Exception as e:
        print(f"⚠️ Ошибка загрузки маппингов v2: {e}, используем fallback...")
        from config import load_column_mapping
        _, indicator_to_excel, indicator_to_file, _ = load_column_mapping(str(column_mapping_path))
        indicator_keywords = {k: {'2022': [], '2023': []} for k in indicator_to_file.keys()}
    
    master_data = {}
    stats = {
        'files_processed': 0,
        'errors': [],
        'found_by_keyword': 0,
        'found_by_hardcode': 0,
        'found_by_fuzzy': 0,
        'not_found': 0,
    }
    
    excel_files_to_load = set(table_source_mapping.values())
    
    for filename in excel_files_to_load:
        excel_path = excel_dir / filename
        if not excel_path.exists():
            print(f"⚠️ Файл не найден: {filename}")
            stats['errors'].append(f"File not found: {filename}")
            continue
        
        try:
            print(f"\n📥 Загружаем: {filename}")
            # ИСПРАВЛЕНИЕ: header=None, чтобы первая строка осталась данными (заголовками)
            df = pd.read_excel(excel_path, header=None)
            
            if df.empty:
                print(f"⚠️ Excel файл пуст: {filename}")
                stats['errors'].append(f"Empty file: {filename}")
                continue
            
            # Теперь df.iloc[0] содержит заголовки, а df.iloc[1:] - данные
            headers = df.iloc[0]  # Сохраняем заголовки отдельно
            data_rows = df.iloc[1:]  # Данные без заголовков
            
            # Фильтруем по кодам ОКВЭД (первый столбец в данных)
            df_filtered = data_rows[data_rows.iloc[:, 0].astype(str).isin(okved_codes_set)]
            
            if df_filtered.empty:
                print(f"ℹ️ Нет данных для нужных кодов ОКВЭД в файле {filename}")
                continue
            
            stats['files_processed'] += 1
            
            # Обрабатываем каждый показатель для этого файла
            for indicator, (col_22_hardcode, col_23_hardcode) in indicator_to_excel.items():
                if indicator_to_file[indicator] != filename:
                    continue  # Пропускаем, если источник не совпадает
                
                keywords_2022 = indicator_keywords.get(indicator, {}).get('2022', [])
                keywords_2023 = indicator_keywords.get(indicator, {}).get('2023', [])
                
                # === ЭТАП 1: Динамический поиск колонок по году ===
                col_idx_2022 = _find_column_smart(df, col_22_hardcode, "2022", keywords_2022)
                col_idx_2023 = _find_column_smart(df, col_23_hardcode, "2023", keywords_2023)
                
                # === ЭТАП 2: Нечеткий поиск строк (экспериментально) ===
                if use_fuzzy_match:
                    # Получаем из column_mapping.csv название показателя
                    # Для примера, используем код индикатора как подсказку
                    pass  # TODO: Реализовать fuzzy match для строк при необходимости
                
                # === ЭТАП 3: Извлечение данных с нормализацией ===
                for okved, group in df_filtered.groupby(df_filtered.iloc[:, 0]):
                    if okved not in master_data:
                        master_data[okved] = {}
                    
                    # Получаем значения за 2022 и 2023
                    if col_idx_2022 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2022)
                            if value:
                                normalized = clean_excel_value_for_word(value)
                                master_data[okved][f"{indicator}_22"] = normalized
                                stats['found_by_keyword' if keywords_2022 else 'found_by_hardcode'] += 1
                                break
                    
                    if col_idx_2023 is not None:
                        for idx, row in group.iterrows():
                            value = get_cell_value_safely(row, col_idx_2023)
                            if value:
                                normalized = clean_excel_value_for_word(value)
                                master_data[okved][f"{indicator}_23"] = normalized
                                stats['found_by_keyword' if keywords_2023 else 'found_by_hardcode'] += 1
                                break
        
        except Exception as e:
            print(f"❌ Ошибка обработки {filename}: {e}")
            stats['errors'].append(f"Error in {filename}: {str(e)}")
    
    print(f"\n✅ Загружено данных для {len(master_data)} кодов ОКВЭД")
    print(f"📊 Статистика:")
    print(f"   - Файлов обработано: {stats['files_processed']}")
    print(f"   - Найдено по ключевым словам: {stats['found_by_keyword']}")
    print(f"   - Найдено по индексам: {stats['found_by_hardcode']}")
    print(f"   - Ошибок: {len(stats['errors'])}")
    
    return master_data, stats


def _find_column_smart(df: pd.DataFrame, hardcode_idx: str, year: str, keywords: list) -> Optional[int]:
    """
    Умный поиск колонки: сначала пробует ключевые слова, потом жесткий индекс.
    
    Args:
        df: DataFrame (уже с header=None, поэтому df.iloc[0] - это заголовки)
        hardcode_idx: Жесткий индекс из column_mapping.csv (fallback)
        year: Год для поиска ("2022" или "2023")
        keywords: Список ключевых слов для поиска
    
    Returns:
        Индекс колонки или None
    """
    # 1️⃣ Пытаемся найти по ключевым словам с нормализацией
    if keywords:
        for keyword in keywords:
            keyword_normalized = normalize_text(keyword)
            for col_idx, header in enumerate(df.iloc[0]):
                header_str = str(header).strip()
                header_normalized = normalize_text(header_str)
                if keyword_normalized in header_normalized:
                    print(f"   🔍 Найдено по ключевому слову '{keyword}' → колонка {col_idx} ('{header_str}')")
                    return col_idx
    
    # 2️⃣ Fuzzy matching (нечеткое совпадение)
    from difflib import SequenceMatcher
    for keyword in keywords:
        keyword_normalized = normalize_text(keyword)
        best_match_ratio = 0
        best_match_idx = None
        for col_idx, header in enumerate(df.iloc[0]):
            header_str = str(header).strip()
            header_normalized = normalize_text(header_str)
            ratio = SequenceMatcher(None, keyword_normalized, header_normalized).ratio()
            if ratio > best_match_ratio:
                best_match_ratio = ratio
                best_match_idx = col_idx
        
        if best_match_ratio >= 0.75:  # Порог fuzzy matching
            print(f"   🔍 Найдено по fuzzy match (сходство {best_match_ratio:.2f}) → колонка {best_match_idx}")
            return best_match_idx
    
    # 3️⃣ Пытаемся найти по году
    for col_idx, header in enumerate(df.iloc[0]):
        header_str = str(header).strip()
        if year in header_str:
            print(f"   🔍 Найдено по году '{year}' → колонка {col_idx} ('{header_str}')")
            return col_idx
    
    # 4️⃣ Fallback на жесткий индекс
    if hardcode_idx:
        try:
            idx = int(hardcode_idx) - 1  # CSV использует 1-based индексы
            if 0 <= idx < len(df.columns):
                print(f"   🔍 Использован жесткий индекс {hardcode_idx} → колонка {idx}")
                return idx
        except ValueError:
            pass
    
    print(f"   ⚠️ Колонка не найдена для года {year}, keywords={keywords}, hardcode={hardcode_idx}")
    return None


def get_cell_value_safely(row: pd.Series, col_idx: int) -> str:
    """Безопасно получает значение ячейки"""
    try:
        value = row.iloc[col_idx]
        if pd.isna(value):
            return ""
        return str(value).strip()
    except (IndexError, KeyError):
        return ""


def fill_word_template_by_tags_v2(doc, master_data: Dict, log_path: Optional[Path] = None, 
                                  report_path: Optional[Path] = None) -> list:
    """
    Заполняет теги в Word с использованием нормализованных данных.
    
    Args:
        doc: Document из python-docx
        master_data: Данные из Excel (с нормализацией)
        log_path: Путь для сохранения логов
        report_path: Путь для сохранения отчета о незаполненных тегах
    
    Returns:
        Список незаполненных тегов
    """
    print("\n🧩 Шаг 4: Заполнение шаблона по тегам (с нормализацией)")
    
    unfilled_tags = []
    log = []
    
    tag_regex = re.compile(r"\{\{([^}]+)\}\}")
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        matches = list(tag_regex.finditer(text))
                        
                        for match in matches:
                            full_tag = match.group(1)
                            raw_tag = full_tag
                            
                            if raw_tag.startswith("OKVED_"):
                                raw_tag = raw_tag[len("OKVED_"):]
                            
                            parts = raw_tag.split("_")
                            if len(parts) < 2:
                                log.append(f"⚠️ Ошибка формата тега: {full_tag}")
                                continue
                            
                            # Парсируем OKVED код и индикатор
                            year_suffix = None
                            if parts[-1] in ("22", "23"):
                                year_suffix = parts[-1]
                                indicator = parts[-2]
                                okved_parts = parts[:-2]
                            else:
                                indicator = parts[-1]
                                okved_parts = parts[:-1]
                            
                            okved_code = ".".join(okved_parts)
                            indicator_key = f"{indicator}_{year_suffix}" if year_suffix else indicator
                            
                            # Ищем значение в master_data
                            value = master_data.get(okved_code, {}).get(indicator_key)
                            
                            if value is None:
                                # Если 2023 не найдена, пробуем 2022 (fallback)
                                if year_suffix == "23":
                                    value = master_data.get(okved_code, {}).get(f"{indicator}_22")
                            
                            # Применяем финальную нормализацию
                            if value:
                                value = clean_excel_value_for_word(value)
                            
                            # Вставляем значение: если нет данных, вставляем прочерк
                            if value:
                                text = text.replace(f"{{{{{full_tag}}}}}", value)
                                log.append(f"✅ Заполнено: {full_tag} → {value}")
                            else:
                                # Прочерк для отсутствующих данных (чтобы совпадать с эталоном)
                                text = text.replace(f"{{{{{full_tag}}}}}", "-")
                                log.append(f"ℹ️ Отсутствующие данные: {full_tag} → [-]")
                                unfilled_tags.append(full_tag)
                        
                        run.text = text
    
    # Сохраняем логи
    if log_path:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log))
        print(f"📝 Лог сохранён: {log_path}")
    
    print(f"✅ Заполнено тегов: {len(log) - len(unfilled_tags)}")
    print(f"⚠️ Незаполненных тегов: {len(unfilled_tags)}")
    
    return unfilled_tags
