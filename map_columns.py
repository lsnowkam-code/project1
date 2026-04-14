# map_columns.py

# Словарь для генерации тегов на основе заголовков в Word
WORD_COLUMNS_MAP = {
    'Валюта баланса': ('ValBal_22', 'ValBal_23'),
    'Внеоборотные активы': ('VneObAk_22', 'VneObAk_23'),
    'Оборотные активы': ('ObAk_22', 'ObAk_23'),
    'Капитал и резервы': ('KapRez_22', 'KapRez_23'),
    'Долгосрочные обязательства': ('DlOb_22', 'DlOb_23'),
    'Краткосрочные обязательства': ('KrOb_22', 'KrOb_23')
}

# Годовые суффиксы (если нужны для других целей)
# должно быть соответствие между колонками в word и excel на конец предыдущего года = 2022, на конец отчетного года = 2023
YEAR_SUFFIXES = [('_22', '2022'), ('_23', '2023')]
YEAR_SUFFIX_MAP = {i: YEAR_SUFFIXES[i] for i in range(len(YEAR_SUFFIXES))}

__all__ = ['WORD_COLUMNS_MAP', 'YEAR_SUFFIXES', 'YEAR_SUFFIX_MAP']