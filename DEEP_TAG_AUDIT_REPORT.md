# Deep audit DOCX (через ZIP/XML, без python-docx)

Дата: 2026-04-23.

## Что проверяли
Проверка таблиц **5, 8, 14, 17** на дубли тегов в строках:
- `output/Бюллетень_17.2.8_ГОТОВЫЙ.docx`
- `output/Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx`

## Результат

### 1) Финальный заполненный документ
Файл: `output/Бюллетень_17.2.8_ГОТОВЫЙ.docx`
- table 5: rows_with_duplicates=0
- table 8: rows_with_duplicates=0
- table 14: rows_with_duplicates=0
- table 17: rows_with_duplicates=0

Интерпретация: в заполненном документе тегов практически нет (они заменены числами), поэтому дубли тегов там не детектируются.

### 2) Шаблон с тегами
Файл: `output/Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx`
- table 5: rows_with_duplicates=56, extra_duplicate_occurrences=224
- table 8: rows_with_duplicates=47, extra_duplicate_occurrences=282
- table 14: rows_with_duplicates=56, extra_duplicate_occurrences=112
- table 17: rows_with_duplicates=47, extra_duplicate_occurrences=106

Интерпретация: проблема дублирования тегов в критичных таблицах **сохраняется в шаблоне**.

## Вывод
Статус: **«проблема осталась»** (в шаблоне с тегами).

