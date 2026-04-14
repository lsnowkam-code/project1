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

        for okved, values_by_col_num in excel_data.items():
            if okved not in master_data:
                master_data[okved] = {}

            for indicator, (col_22, col_23) in indicator_to_excel.items():
                if indicator_to_file[indicator] != filename:
                    continue  # ❗ Пропускаем, если источник не совпадает

                if str(col_22) in values_by_col_num:
                    master_data[okved][f"{indicator}_22"] = values_by_col_num[str(col_22)]
                if str(col_23) in values_by_col_num:
                    master_data[okved][f"{indicator}_23"] = values_by_col_num[str(col_23)]

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
    unfilled_tags = []
    log = []

    # Собираем список допустимых тегов из column_mapping
    valid_indicators = set(column_mapping["Код показателя"].astype(str))

    # Обрабатываем только таблицы
    for table in doc.tables:
        for row in table.rows:
            row_text = " ".join(cell.text for cell in row.cells)
            if not any(code in row_text for code in master_data.keys()):
                continue  # пропускаем строки без ОКВЭД или муниципалитета

            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        matches = extract_tags(run.text)
                        for tag in matches:
                            if "_" not in tag:
                                continue
                            indicator, code = tag.split("_", 1)
                            indicator = indicator.strip()
                            code = code.strip()

                            if indicator not in valid_indicators:
                                log.append(f"⏭️ Пропуск: тег '{tag}' не найден в маппинге")
                                continue

                            if code not in master_data:
                                log.append(f"⏭️ Пропуск: код '{code}' не найден в master_data")
                                continue

                            value = master_data[code].get(indicator)
                            if value is None or pd.isna(value):
                                run.text = replace_tag(run.text, tag, "—")
                                log.append(f"⚠️ Нет данных: {tag} → '—'")
                                unfilled_tags.append({"tag": tag, "indicator": indicator, "code": code})
                            else:
                                run.text = replace_tag(run.text, tag, str(value))
                                log.append(f"✅ Подставлено: {tag} → {value}")

    # Сохраняем лог
    if log_path:
        with open(log_path, "w", encoding="utf-8") as f:
            for line in log:
                f.write(line + "\n")

    # Сохраняем незаполненные теги в Excel
    if report_path and unfilled_tags:
        pd.DataFrame(unfilled_tags).to_excel(report_path, index=False)

    return [t["tag"] for t in unfilled_tags]
