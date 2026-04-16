# map_columns.py

# ⚠️ ВНИМАНИЕ: Этот файл больше не используется в основной логике!
# WORD_COLUMNS_MAP был заменен на использование column_mapping.csv
# YEAR_SUFFIXES и YEAR_SUFFIX_MAP больше не нужны

# Все мапинги теперь загружаются из CSV файлов:
# - column_mapping.csv для соответствия названий показателей
# - table_source_data_mapping.csv для соответствия таблиц и файлов

# Этот файл оставлен для совместимости со старыми файлами в input/excel/
# но в основной логике проекта не используется

# УСТАРЕВШИЕ ПЕРЕМЕННЫЕ (не использовать в новом коде):
WORD_COLUMNS_MAP = {
    'Валюта баланса': ('ValBal_22', 'ValBal_23'),
    'Внеоборотные активы': ('VneObAk_22', 'VneObAk_23'),
    'Оборотные активы': ('ObAk_22', 'ObAk_23'),
    'Капитал и резервы': ('KapRez_22', 'KapRez_23'),
    'Долгосрочные обязательства': ('DlOb_22', 'DlOb_23'),
    'Краткосрочные обязательства': ('KrOb_22', 'KrOb_23')
}

YEAR_SUFFIXES = [('_22', '2022'), ('_23', '2023')]
YEAR_SUFFIX_MAP = {i: YEAR_SUFFIXES[i] for i in range(len(YEAR_SUFFIXES))}

# Экспортируем только для совместимости со старыми файлами
__all__ = ['WORD_COLUMNS_MAP', 'YEAR_SUFFIXES', 'YEAR_SUFFIX_MAP']