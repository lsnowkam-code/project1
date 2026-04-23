# ⚡ QUICK START - Запуск v2.0

## 🏃 За 5 минут

### Шаг 1: Переверьте тесты (2 минуты)
```bash
cd /workspaces/project
pip install -r requirements.txt
python test_smart_loader.py
```

Ожидайте вывод:
```
🧪 ЗАПУСК ТЕСТОВ УМНОГО ПОИСКА
═══════════════════════════════════════════════════
✅ Тест 1: Нормализация текста
✅ Тест 2: Нечеткое совпадение (fuzzy_match)
✅ Тест 3: Динамический поиск колонок
✅ Тест 4: Нечеткий поиск строк
✅ Тест 5: Нормализация чисел
✅ Тест 6: Комплексная очистка значений
✅ ТЕСТЫ ЗАВЕРШЕНЫ
═══════════════════════════════════════════════════
```

**Если все ✅ - переходите к шагу 2**

---

### Шаг 2: Обновляем main.py (3 минуты)

Откройте основной файл и найдите строку:
```python
from data_filler import pre_load_all_excel_data, fill_word_template_by_tags
```

Замените на:
```python
from data_filler_v2 import pre_load_all_excel_data_v2, fill_word_template_by_tags_v2
from config_v2 import load_column_mapping_v2
```

Найдите где вызывается:
```python
master_data = pre_load_all_excel_data(excel_dir, table_source_mapping, okved_codes_set, column_mapping_path)
```

Замените на:
```python
master_data, stats = pre_load_all_excel_data_v2(
    excel_dir=excel_dir,
    table_source_mapping=table_source_mapping,
    okved_codes_set=okved_codes_set,
    column_mapping_path=column_mapping_path,
    use_fuzzy_match=True,
    fuzzy_threshold=0.80
)
print(f"📊 Статистика загрузки: {stats}")
```

Найдите:
```python
fill_word_template_by_tags(doc, master_data, column_mapping, log_path=log_path)
```

Замените на:
```python
fill_word_template_by_tags_v2(doc, master_data, log_path=log_path)
```

**Done! ✅**

---

### Шаг 3: Запустите!

```bash
python main.py
```

**Если все работает:**
```
📊 Шаг 3: Загрузка данных из Excel (УМНЫЙ поиск)
📥 Загружаем: T23_000000_t01Ved14.xlsx
📊 Статистика загрузки: {
  'files_processed': 7,
  'found_by_keyword': 42,
  'found_by_hardcode': 8,
  'errors': []
}
🧩 Шаг 4: Заполнение шаблона по тегам (с нормализацией)
✅ Заполнено тегов: 280
⚠️ Незаполненных: 5
```

---

## 🎯 Результат

Сравните файлы:
```bash
# Windows:
diff Бюллетень_17.2.8_ГОТОВЫЙ.docx "как должно быть!.docx"

# Linux/Mac:
cmp Бюллетень_17.2.8_ГОТОВЫЙ.docx "как должно быть!.docx"
```

**Числа должны совпадать:**
- ✅ "569800589" (без пробелов)
- ✅ "-" (минус правильный)
- ✅ "2023" (за правильный год)

---

## 🔙 Если что-то сломалось

### Откатиться на v1:
```python
# В main.py измените обратно на:
from data_filler import pre_load_all_excel_data, fill_word_template_by_tags

# И вызовите как было:
master_data = pre_load_all_excel_data(...)
fill_word_template_by_tags(doc, master_data, ...)
```

### Параллельный запуск:
Если хотите тестировать v1 и v2 одновременно:

```python
# В main.py добавьте параметр:
USE_V2 = True  # Переключение между версиями

if USE_V2:
    from data_filler_v2 import pre_load_all_excel_data_v2, fill_word_template_by_tags_v2
    master_data, stats = pre_load_all_excel_data_v2(...)
    fill_word_template_by_tags_v2(doc, master_data, ...)
else:
    from data_filler import pre_load_all_excel_data, fill_word_template_by_tags
    master_data = pre_load_all_excel_data(...)
    fill_word_template_by_tags(doc, master_data, ...)
```

---

## 📊 Что получилось

### Файлы добавлены:
```
✅ smart_loader.py          - Умный поиск (340 строк)
✅ data_normalizer.py       - Очистка данных (160 строк)
✅ config_v2.py             - Расширенный конфиг (180 строк)
✅ data_filler_v2.py        - Новый заполнитель (280 строк)
✅ test_smart_loader.py     - Тесты (250 строк)
✅ column_mapping_v2.csv    - Маппинги с ключевыми словами
```

### Маппинги обновлены:
```csv
Название показателя,Код показателя,...,Ключевое слово 2022,Ключевое слово 2023
Валюта баланса,ValBal,...,"2022","На начало","2023","На конец"
```

---

## 📖 Полная документация

- **IMPLEMENTATION_PLAN_V2.md** - Пошаговое руководство
- **SUMMARY_V2.md** - Полная сводка изменений
- **QUICK_START_V2.md** - Быстрый старт

---

## ✨ Преимущества v2

| Функция | v1 | v2 |
|---------|----|----|
| Жесткие индексы | ❌ | ✅ Ключевые слова |
| Устойчива к опечаткам | ❌ | ✅ Fuzzy match |
| Автоматическая нормализация | ❌ | ✅ 100% совпадение |
| Логирование | Базовое | ✅ Подробное |
| Статистика | ❌ | ✅ Полная |

---

**🚀 Готово! Вперед к более надежной системе!**
