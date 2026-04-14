# data_filler.py

import re
from config import get_excel_data, load_column_mapping
import pandas as pd


def pre_load_all_excel_data(excel_dir, table_source_mapping, okved_codes_set, column_mapping_path):
    """
    Загружает все данные из Excel по справочнику столбцов.
    Возвращает master_data[okved][indicator_year] = value
    """
    print("\n📊 Шаг 3: Загрузка данных из Excel")

    _, indicator_to_excel, indicator_to_file = load_column_mapping(column_mapping_path)
    master_data = {}

    excel_files_to_load = set(table_source_mapping.values())

    for filename in excel_files_to_load:
        excel_path = excel_dir / filename
        if not excel_path.exists():
            print(f"⚠️ Файл не найден: {filename}")
            continue

        print(f"📥 Загружаем: {filename}")
        excel_data = get_excel_data(excel_path, okved_codes_set)

        # 👇 Лог по каждому показателю
        indicators_loaded = []

        for okved, values_by_col_num in excel_data.items():
            if okved not in master_data:
                master_data[okved] = {}

            for indicator, (col_22, col_23) in indicator_to_excel.items():
                if indicator_to_file[indicator] != filename:
                    continue  # ❗ Пропускаем, если источник не совпадает

                loaded = False

                if str(col_22) in values_by_col_num:
                    master_data[okved][f"{indicator}_22"] = values_by_col_num[str(col_22)]
                    loaded = True

                if str(col_23) in values_by_col_num:
                    master_data[okved][f"{indicator}_23"] = values_by_col_num[str(col_23)]
                    loaded = True

                if loaded:
                    indicators_loaded.append(indicator)

        # 📊 Сводка по файлу
        if indicators_loaded:
            unique_indicators = sorted(set(indicators_loaded))
            print(f"   ✅ Показатели из {filename}: {', '.join(unique_indicators)}")
        else:
            print(f"   ⚠️ Нет показателей, соответствующих этому файлу")

    print(f"✅ Загружено данных для {len(master_data)} кодов ОКВЭД")
    return master_data


"""def fill_word_template_by_tags(doc, master_data):
    Заполняет шаблон Word по тегам {{OKVED_INDICATOR_YEAR}}.
    print("\n🧩 Шаг 4: Заполнение шаблона по тегам")
    tag_regex = re.compile(r"\{\{([\w\.]+)_([A-Za-z]+_\d{2})\}\}")

    filled_count = 0
    unfilled_tags = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for match in tag_regex.finditer(paragraph.text):
                        full_tag = match.group(0)
                        okved_code = match.group(1)
                        indicator_key = match.group(2)

                        value = master_data.get(okved_code, {}).get(indicator_key, "—")

                        for run in paragraph.runs:
                            if full_tag in run.text:
                                run.text = run.text.replace(full_tag, str(value))
                                filled_count += 1
                                break

                        if value == "—":
                            unfilled_tags.append(full_tag)

    print(f"✅ Заполнено ячеек: {filled_count}")
    if unfilled_tags:
        print(f"⚠️ Не найдены данные для {len(unfilled_tags)} тегов. Они заменены на '—'.")
    return unfilled_tags"""

TAG_PATTERN = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")


def extract_tags(text):
    return TAG_PATTERN.findall(text)


def replace_tag(text, tag, value):
    return text.replace(f"{{{{{tag}}}}}", str(value))


def fill_word_template_by_tags(doc, master_data, column_mapping, log_path=None, report_path=None):
    """
    Заполняет теги вида {{OKVED_code_indicator_year}}.
    master_data[okved_code][indicator_year] = value
    """
    unfilled_tags = []
    log = []
    
    # Регулярное выражение для поиска тегов: {{OKVED_xxx_YyYy_nn}}
    # Пример: {{OKVED_50_13_11_ValBal_22}}
    tag_regex = re.compile(r"\{\{(OKVED[_A-Za-z0-9]+)\}\}")

    # Обрабатываем только таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        matches = list(tag_regex.finditer(text))
                        
                        for match in matches:
                            full_tag = match.group(1)  # "OKVED_50_13_11_ValBal_22"
                            
                            # Парсим тег: OKVED_<code>_<indicator>_<year>
                            # Пример: OKVED_50_13_11_ValBal_22
                            parts = full_tag.split("_")
                            
                            if len(parts) < 3:
                                log.append(f"⚠️ Ошибка формата тега: {full_tag}")
                                continue
                            
                            # Часть с кодом ОКВЭД (может быть несколько частей через _)
                            # Берем элементы до последних двух (indicator и year)
                            year_suffix = parts[-1]  # "22" или "23"
                            indicator = parts[-2]    # "ValBal"
                            okved_parts = parts[1:-2]  # ["50", "13", "11"]
                            okved_code = ".".join(okved_parts)  # "50.13.11"
                            
                            indicator_year = f"{indicator}_{year_suffix}"  # "ValBal_22"
                            
                            if okved_code not in master_data:
                                log.append(f"⏭️ Код ОКВЭД '{okved_code}' не найден в данных")
                                unfilled_tags.append({"tag": full_tag, "okved_code": okved_code, "indicator": indicator_year, "reason": "Code not found"})
                                continue
                            
                            value = master_data[okved_code].get(indicator_year)
                            if value is None or value == "—":
                                text = text.replace(f"{{{{{full_tag}}}}}", "—")
                                log.append(f"⚠️ Нет данных: {full_tag} ({okved_code}/{indicator_year}) → '—'")
                                unfilled_tags.append({"tag": full_tag, "okved_code": okved_code, "indicator": indicator_year, "reason": "No data"})
                            else:
                                text = text.replace(f"{{{{{full_tag}}}}}", str(value))
                                log.append(f"✅ Заполнено: {full_tag} ({okved_code}/{indicator_year}) → {value}")
                        
                        run.text = text

    # Сохраняем лог
    if log_path:
        with open(log_path, "w", encoding="utf-8") as f:
            for line in log:
                f.write(line + "\n")
        print(f"📝 Лог сохранён: {log_path}")

    # Сохраняем незаполненные теги в Excel
    if report_path and unfilled_tags:
        pd.DataFrame(unfilled_tags).to_excel(report_path, index=False)
        print(f"📊 Отчёт о незаполненных тегах: {report_path}")

    return [t["tag"] for t in unfilled_tags]
