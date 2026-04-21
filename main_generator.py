# main_generator.py
import csv
from docx import Document
from word_parser import parse_word_template, normalize_text
from smart_excel_finder import find_value

# 1. Загрузка маппинга показателей (как в старой системе)
def load_indicator_mapping(mapping_file):
    mapping = {}
    with open(mapping_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            indicator_name = row['Название показателя'].lower()
            excel_file = row['Excel файл']
            mapping[indicator_name] = excel_file
    return mapping

def run_generation(word_template_path, word_output_path, excel_dir, mapping_file):
    # Загружаем маппинг показателей
    indicator_map = load_indicator_mapping(mapping_file)
    
    # Парсим Word шаблон
    tables_structure = parse_word_template(word_template_path)
    doc = Document(word_template_path)
    
    print(f"Найдено таблиц для обработки: {len(tables_structure)}")

    for tbl_info in tables_structure:
        print(f"Проверяем таблицу: {tbl_info['title']}")
        table_title = tbl_info['title'].lower()
        
        # Определяем показатель из заголовка таблицы
        indicator = None
        if 'валют' in table_title and 'баланс' in table_title:
            indicator = "валюта баланса"
        elif 'внеоборотн' in table_title and 'актив' in table_title:
            indicator = "внеоборотные активы"
        elif 'оборотн' in table_title and 'актив' in table_title:
            indicator = "оборотные активы"
        elif 'капитал' in table_title and 'резерв' in table_title:
            indicator = "капитал и резервы"
        elif 'долгосрочн' in table_title and 'обязательств' in table_title:
            indicator = "долгосрочные обязательства"
        elif 'краткосрочн' in table_title and 'обязательств' in table_title:
            indicator = "краткосрочные обязательства"
        else:
            print(f"   ⚠️  Не распознан показатель для таблицы: {tbl_info['title']}")
            continue
            
        # Определяем Excel файл для этого показателя
        indicator_norm = normalize_text(indicator)
        excel_filename = indicator_map.get(indicator_norm)
        
        if not excel_filename:
            print(f"   ⚠️  Не найден маппинг для показателя: {indicator} ({indicator_norm})")
            continue
            
        print(f"📊 Обработка таблицы '{tbl_info['title']}' -> Показатель: {indicator} -> Файл: {excel_filename}")
        
        full_excel_path = f"{excel_dir}/{excel_filename}"

        # Проходим по каждой строке-показателю в таблице Word
        for row_data in tbl_info['rows']:
            okved_norm = row_data['indicator_norm']  # OKVED код
            
            # Определяем показатель из заголовка таблицы (упрощенная логика)
            # В будущем можно улучшить распознавание показателей
            table_title_lower = tbl_info['title'].lower()
            if 'валют' in table_title_lower and 'баланс' in table_title_lower:
                indicator = "валюта баланса"
            elif 'внеоборотн' in table_title_lower and 'актив' in table_title_lower:
                indicator = "внеоборотные активы"
            elif 'оборотн' in table_title_lower and 'актив' in table_title_lower:
                indicator = "оборотные активы"
            elif 'капитал' in table_title_lower and 'резерв' in table_title_lower:
                indicator = "капитал и резервы"
            elif 'долгосрочн' in table_title_lower and 'обязательств' in table_title_lower:
                indicator = "долгосрочные обязательства"
            elif 'краткосрочн' in table_title_lower and 'обязательств' in table_title_lower:
                indicator = "краткосрочные обязательства"
            else:
                indicator = "неизвестный показатель"
                
            indicator_norm = normalize_text(indicator)
            
            for year, word_col_idx in row_data['target_cols'].items():
                # Ищем значение в Excel
                value = find_value(full_excel_path, okved_norm, indicator_norm, year)
                
                if value is not None:
                    # Форматируем значение
                    str_val = str(value)
                    # Удаление пробелов в числах
                    if str_val.replace(' ', '').replace('-', '').replace('.', '').isdigit():
                        str_val = str_val.replace(' ', '')
                    
                    # Создаем тег для замены
                    tag = f"{{{{OKVED_{okved_norm}_{indicator_norm}_{year}}}}}"
                    
                    # Ищем и заменяем тег в документе
                    table = doc.tables[tbl_info['table_index']]
                    for row in table.rows:
                        for cell in row.cells:
                            if tag in cell.text:
                                cell.text = cell.text.replace(tag, str_val)
                                print(f"   ✅ OKVED '{okved_norm}' ({year}): {str_val}")
                                break
                else:
                    print(f"   ❌ Не найдено данных для: OKVED '{okved_norm}' ({year})")

    # Сохраняем результат
    doc.save(word_output_path)
    print(f"✅ Документ сохранен: {word_output_path}")

# Пример запуска
if __name__ == "__main__":
    run_generation(
        word_template_path="output/Бюллетень_ШАБЛОН_С_ТЕГАМИ_v2.docx",
        word_output_path="Бюллетень_17.2.8_ГОТОВЫЙ_ДИНАМИЧЕСКИЙ.docx",
        excel_dir="./input/excel",
        mapping_file="./input/mappings/column_mapping_v2.csv"
    )