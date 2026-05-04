# 🔧 ДЕТАЛЬНЫЙ ПЛАН УСТРАНЕНИЯ КРИТИЧЕСКИХ ПРОБЛЕМ

## 📊 АНАЛИЗ ТЕКУЩЕЙ СИТУАЦИИ

### Статистика проблем:
- **223 незаполненных тега** из 1853总 тегов (12%)
- **Основные проблемные ОКВЭД**: 85, 43, O, 84, 81, 74, 69, 77, Q, 86, 49, 41, 32
- **Проблемные показатели**: VneObAk, DlOb, KrOb, NemAk, OsnSr
- **Дубликаты столбцов**: множественные предупреждения в логе

---

## 🎯 ПРИОРИТЕТ 1: Нормализация кодов ОКВЭД (КРИТИЧНО)

### Проблема:
Теги генерируются в формате: `OKVED_85_VneObAk_23`, `OKVED_R_DlOb_22`
Но в Excel данные хранятся под кодами: `85`, `R`, `101.АГ`, `A`, `B`, `C`...

**Несоответствие возникает потому что:**
1. В шаблоне Word используются коды разделов (A, B, C...) и укрупненные коды (85, 43...)
2. В Excel файлах коды могут быть с точками (101.АГ) или без
3. Функция `find_okved_code()` находит код, но при создании тега используется `.replace('.', '_')`
4. При поиске в `master_data` используется тот же формат, но данные загружены под другими ключами

### Решение:

#### Шаг 1.1: Создать универсальную функцию нормализации ОКВЭД
**Файл:** `config.py` (добавить после строки 286)

```python
def normalize_okved_code_for_tag(okved_code: str) -> str:
    """
    Приводит код ОКВЭД к единому формату для тегов.
    
    Правила:
    - Разделы A-Z остаются как есть: 'A' → 'A'
    - Коды с точкой: '101.АГ' → '101_АГ'
    - Числовые коды: '85', '43' остаются как есть
    - Удаляем лишние пробелы и приводим к верхнему регистру для букв
    """
    if not okved_code:
        return None
    
    code = okved_code.strip()
    
    # Если это раздел (одна буква A-Z)
    if len(code) == 1 and code.upper() in 'ABCDEFGHIJKLMNQPRSUVXZ':
        return code.upper()
    
    # Если есть точка, заменяем на подчеркивание
    if '.' in code:
        return code.replace('.', '_').upper()
    
    # Числовой код - оставляем как есть
    return code


def normalize_okved_code_for_lookup(okved_code: str) -> str:
    """
    Приводит код ОКВЭД к формату для поиска в Excel.
    
    Обратная операция: ищет варианты кода в разных форматах.
    """
    if not okved_code:
        return None
    
    code = okved_code.strip().upper()
    
    # Создаем все возможные варианты для поиска
    variants = [code]
    
    # Если код с подчеркиванием, добавляем вариант с точкой
    if '_' in code:
        variants.append(code.replace('_', '.'))
    
    # Если код без точки, добавляем возможные форматы
    if code.isdigit():
        # Для числовых кодов добавляем варианты
        variants.append(code)
        if len(code) == 2:
            variants.append(f"{code}.0")  # На случай если в Excel "85.0"
    
    return variants
```

#### Шаг 1.2: Модифицировать `generate_word_template()` в `logic.py`
**Строка 344:** Заменить текущую логику формирования `okved_tag_part`

```python
# БЫЛО (строка 344):
# okved_tag_part = okved_code.replace('.', '_')

# СТАЛО:
from config import normalize_okved_code_for_tag
okved_tag_part = normalize_okved_code_for_tag(okved_code)
if not okved_tag_part:
    print(f"   ⚠️ Не удалось нормализовать код ОКВЭД: {okved_code}")
    continue
```

#### Шаг 1.3: Модифицировать загрузку Excel данных
**Файл:** `data_filler_v2.py`, функция `pre_load_all_excel_data_v2()`

После строки 81 (где фильтруется df_filtered):
```python
# ДОБАВИТЬ: Создание обратного маппинга кодов
from config import normalize_okved_code_for_tag

# Создаем маппинг: нормализованный код → оригинальный код
okved_reverse_map = {}
for code in okved_codes_set:
    normalized = normalize_okved_code_for_tag(code)
    if normalized:
        okved_reverse_map[normalized] = code

# ... внутри цикла по OKVED (после строки 109):
for okved, group in df_filtered.groupby(df_filtered.iloc[:, 0]):
    # Добавляем запись и по нормализованному коду
    normalized_okved = normalize_okved_code_for_tag(okved)
    
    # Сохраняем данные под ОБОИМИ ключами
    target_keys = [okved]
    if normalized_okved and normalized_okved != okved:
        target_keys.append(normalized_okved)
    
    for target_key in target_keys:
        if target_key not in master_data:
            master_data[target_key] = {}
        
        # ... остальная логика заполнения
```

#### Шаг 1.4: Модифицировать парсинг тегов в `fill_word_template_by_tags_v2()`
**Строки 254-258:** Улучшить поиск по OKVED коду

```python
# БЫЛО:
# okved_code = ".".join(okved_parts)
# indicator_key = f"{indicator}_{year_suffix}" if year_suffix else indicator
# value = master_data.get(okved_code, {}).get(indicator_key)

# СТАЛО:
from config import normalize_okved_code_for_lookup

# Восстанавливаем возможный формат кода
okved_base = "_".join(okved_parts)
possible_codes = normalize_okved_code_for_lookup(okved_base)

# Ищем по всем возможным вариантам
value = None
if possible_codes:
    for code_variant in possible_codes:
        if code_variant in master_data:
            value = master_data[code_variant].get(indicator_key)
            if value:
                break

# Fallback: пробуем прямой поиск
if value is None:
    okved_direct = okved_base.replace('_', '.')
    value = master_data.get(okved_direct, {}).get(indicator_key)
```

---

## 🎯 ПРИОРИТЕТ 2: Устранение дубликатов столбцов

### Проблема:
Функция `_find_header_rows()` находит слишком много "заголовков", включая строки с данными.
Предупреждения: "⚠️ Пропускаем дубликат столбца X для..."

### Решение:

#### Шаг 2.1: Улучшить `_find_header_rows()` в `logic.py`
**Строки 48-58:** Добавить более строгую фильтрацию

```python
def _find_header_rows(table, source_word_to_indicator, max_search_rows=80):
    """Собирает все строки заголовков таблицы (year или indicator rows)."""
    headers = []
    
    # Считаем количество лет в каждой строке
    for row_idx, row in enumerate(table.rows[:max_search_rows]):
        row_years = [get_cleaned_cell_text(cell).strip() for cell in row.cells]
        year_count = sum(1 for year in row_years if year in ('2022', '2023'))
        
        # Строка с годами - точно заголовок
        if year_count >= 2:
            headers.append(row_idx)
            continue
        
        # Проверяем на индикаторы
        row_texts = [_normalize_text(get_cleaned_cell_text(cell)) for cell in row.cells]
        indicator_score = sum(
            1 for cell_text in row_texts 
            for name in source_word_to_indicator.keys() 
            if name in cell_text and len(cell_text) > 5  # Игнорируем короткие совпадения
        )
        
        # Заголовок должен содержать минимум 2 индикатора ИЛИ быть очень похожим
        if indicator_score >= 2:
            # Дополнительная проверка: не содержит ли строка чисел (признак данных)
            has_numbers = any(cell[0].isdigit() for cell in row_texts if cell)
            if not has_numbers:
                headers.append(row_idx)
    
    # Удаляем дубликаты и близкие строки
    if len(headers) > 1:
        filtered_headers = []
        for i, h in enumerate(sorted(set(headers))):
            # Если заголовки идут подряд (разница < 2), берем последний
            if i == 0 or h - filtered_headers[-1] >= 2:
                filtered_headers.append(h)
            else:
                filtered_headers[-1] = h  # Заменяем на более поздний
        headers = filtered_headers
    
    return headers
```

#### Шаг 2.2: Улучшить `_compute_section_mapping()` для обработки дубликатов
**Строки 134-145:** Логировать причину дублирования

```python
# Убираем дубликаты с расширенным логированием
seen_specs = set()
filtered_map = {}
duplicate_info = {}

for col_idx in sorted(col_to_indicator_map):
    spec = col_to_indicator_map[col_idx]
    if spec in seen_specs:
        # Логируем первый столбец с таким spec
        first_col = duplicate_info.get(spec, 'unknown')
        print(f"   ⚠️ Дубликат: столбец {col_idx} ({spec[0]}_{spec[1] if spec[1] else ''}) "
              f"- дублирует столбец {first_col}")
        continue
    seen_specs.add(spec)
    duplicate_info[spec] = col_idx
    filtered_map[col_idx] = spec

return filtered_map
```

---

## 🎯 ПРИОРИТЕТ 3: Улучшение определения источников таблиц

### Проблема:
- Таблица 17 жестко задана (хрупкое решение)
- Функция `get_table_name()` может не найти заголовок

### Решение:

#### Шаг 3.1: Удалить хардкод для таблицы 17
**Строки 210-214 в `logic.py`:** Заменить на автоматическое определение

```python
# БЫЛО:
# if t_index + 1 == 17:
#     print(f"⚠️ Таблица 17 - применяем специальное определение (таблица 7)")
#     current_source_file = 'T23_000000_t20Ved14.xlsx'

# СТАЛО:
# Определяем по метке "Продолжение таблицы 17" или по содержимому
continuation_number = get_continuation_table_number(table)
if continuation_number == 17:
    # Таблица 17 соответствует таблице 7 в маппинге
    table_source_list = list(table_source_mapping.values())
    if len(table_source_list) >= 7:
        current_source_file = table_source_list[6]  # Индекс 6 = таблица 7
        print(f"🔁 Таблица 17 → источник таблицы 7: {current_source_file}")
```

#### Шаг 3.2: Улучшить `auto_detect_table_source()`
**Строки 151-167 в `logic.py`:** Расширить поиск

```python
def auto_detect_table_source(table, table_source_mapping, file_word_to_indicator):
    """Автоматически определяет источник Excel файла для таблицы."""
    # Собираем текст из первых 10 строк таблицы (было 5)
    table_content = ' '.join([
        get_cleaned_cell_text(cell).lower()
        for row in table.rows[:10]
        for cell in row.cells
    ])
    
    # Считаем совпадения по каждому файлу
    file_scores = {}
    for (excel_file, indicator_name), indicator in file_word_to_indicator.items():
        if indicator_name.lower() in table_content:
            file_scores[excel_file] = file_scores.get(excel_file, 0) + 1
    
    if file_scores:
        best_file = max(file_scores, key=file_scores.get)
        print(f"   ✓ Автоматически определен источник: {best_file} "
              f"(совпадений: {file_scores[best_file]})")
        return best_file
    
    return None
```

---

## 🎯 ПРИОРИТЕТ 4: Валидация и диагностика

### Шаг 4.1: Добавить предобработку данных с валидацией
**Новый файл:** `validation.py`

```python
def validate_okved_coverage(master_data, template_tags):
    """
    Проверяет, какие теги не могут быть заполнены из-за отсутствия данных.
    
    Returns: dict с статистикой и рекомендациями
    """
    from collections import defaultdict
    
    missing_by_okved = defaultdict(list)
    missing_by_indicator = defaultdict(list)
    
    for tag in template_tags:
        parts = tag.replace('OKVED_', '').split('_')
        if len(parts) < 2:
            continue
        
        okved = parts[0]
        indicator = parts[-2] if parts[-1] in ('22', '23') else parts[-1]
        
        if okved not in master_data or indicator not in str(master_data.get(okved, {})):
            missing_by_okved[okved].append(tag)
            missing_by_indicator[indicator].append(tag)
    
    return {
        'total_missing': len(template_tags),
        'by_okved': dict(missing_by_okved),
        'by_indicator': dict(missing_by_indicator),
        'top_missing_okved': sorted(missing_by_okved.items(), key=lambda x: -len(x[1]))[:10],
        'top_missing_indicators': sorted(missing_by_indicator.items(), key=lambda x: -len(x[1]))[:10]
    }
```

### Шаг 4.2: Добавить отчет о покрытии данных
**В конец `main.py`:**

```python
# После заполнения шаблона
from validation import validate_okved_coverage

# Собираем все теги из шаблона
all_tags = set()
tag_regex = re.compile(r"\{\{([^}]+)\}\}")
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                matches = tag_regex.findall(paragraph.text)
                all_tags.update(matches)

# Проверяем покрытие
report = validate_okved_coverage(master_data, all_tags)
print("\n📊 ОТЧЕТ О ПОКРЫТИИ ДАННЫХ:")
print(f"   Всего тегов: {report['total_missing']}")
print(f"   Топ проблемных ОКВЭД:")
for okved, tags in report['top_missing_okved']:
    print(f"      {okved}: {len(tags)} тегов")
print(f"   Топ проблемных показателей:")
for ind, tags in report['top_missing_indicators']:
    print(f"      {ind}: {len(tags)} тегов")
```

---

## 📅 ПЛАН РЕАЛИЗАЦИИ ПО ЭТАПАМ

### Этап 1: Нормализация ОКВЭД (4-6 часов)
1. ✅ Добавить функции нормализации в `config.py`
2. ✅ Модифицировать `logic.py` (генерация тегов)
3. ✅ Модифицировать `data_filler_v2.py` (загрузка Excel)
4. ✅ Модифицировать `data_filler_v2.py` (парсинг тегов)
5. ⏳ Протестировать на текущих данных

### Этап 2: Устранение дубликатов (2-3 часа)
1. ✅ Улучшить `_find_header_rows()`
2. ✅ Добавить расширенное логирование дубликатов
3. ⏳ Проверить логи на предмет дубликатов

### Этап 3: Улучшение определения источников (1-2 часа)
1. ✅ Удалить хардкод таблицы 17
2. ✅ Улучшить `auto_detect_table_source()`
3. ⏳ Протестировать на всех таблицах

### Этап 4: Валидация (2-3 часа)
1. ✅ Создать `validation.py`
2. ✅ Добавить отчет в `main.py`
3. ⏳ Сгенерировать финальный отчет

---

## ✅ КРИТЕРИИ УСПЕХА

1. **Незаполненные теги**: ≤ 20 (сейчас 223)
2. **Дубликаты столбцов**: ≤ 5 предупреждений (сейчас множественные)
3. **Определение источников**: 100% таблиц определены автоматически
4. **Покрытие ОКВЭД**: ≥ 95% кодов из шаблона есть в Excel

---

## 🔍 МОНИТОРИНГ ПОСЛЕ ИСПРАВЛЕНИЙ

Запустить команду и проверить:
```bash
grep "ℹ️ Отсутствующие" output/fill_log.txt | wc -l
# Должно быть ≤ 20

grep "⚠️ Пропускаем дубликат" output/*.log | wc -l
# Должно быть ≤ 5
```
