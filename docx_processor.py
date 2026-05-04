import re
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

# --- УТИЛИТЫ ---

def normalize(text):
    """Нормализует текст: убирает лишние пробелы, приводит к нижнему регистру."""
    if not text:
        return ""
    return " ".join(text.split()).strip().lower()


def load_mapping(path):
    """Загружает названия таблиц из CSV (разделитель ;)."""
    import pandas as pd
    try:
        df = pd.read_csv(path, sep=";")
        # Предполагаем, что колонка называется "Таблица"
        if "Таблица" not in df.columns:
            raise ValueError("В mapping.csv отсутствует колонка 'Таблица'")
        return [normalize(x) for x in df["Таблица"].tolist()]
    except FileNotFoundError:
        print(f"[ERROR] Файл {path} не найден.")
        return []
    except Exception as e:
        print(f"[ERROR] Ошибка при чтении mapping.csv: {e}")
        return []


# --- ИТЕРАЦИЯ ПО ДОКУМЕНТУ ---

def iter_block_items(doc):
    """
    Идём по документу строго в порядке следования элементов:
    paragraph → table → paragraph → table
    Возвращает объекты docx.Paragraph или docx.Table.
    """
    body = doc._element.body

    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


# --- ДЕТЕКЦИЯ БЛОКОВ ---

def detect_blocks(doc, mapping_names):
    """
    Находит блоки документа.
    Блок начинается с параграфа, текст которого есть в mapping_names.
    Все таблицы, идущие после заголовка до следующего заголовка, принадлежат блоку.
    """
    blocks = []
    current_block = None

    for item in iter_block_items(doc):
        # --- PARAGRAPH ---
        if isinstance(item, Paragraph):
            text = normalize(item.text)

            # Игнорируем пустые строки
            if not text:
                continue

            if text in mapping_names:
                print(f"[BLOCK START] {text}")
                current_block = {
                    "name": text,
                    "tables": []
                }
                blocks.append(current_block)
            # Если это не заголовок блока, просто игнорируем параграф
            # (текст между таблицами внутри блока нам пока не важен)

        # --- TABLE ---
        elif isinstance(item, Table):
            if current_block is None:
                # Таблица до первого найденного заголовка
                print("[WARN] Table found outside any block. Skipping.")
                continue

            current_block["tables"].append(item)

    return blocks


# --- ИЗВЛЕЧЕНИЕ СХЕМЫ (СТРУКТУРЫ) ---

def extract_year(text):
    """Ищет год вида 20XX в тексте."""
    match = re.search(r"20\d{2}", text)
    return match.group(0) if match else None


def remove_year(text):
    """Удаляет год вида 20XX из текста."""
    return re.sub(r"20\d{2}", "", text).strip()


def extract_columns(table):
    """
    Извлекает заголовки столбцов из первой строки таблицы.
    Возвращает список словарей: {'name': '...', 'year': '...'}
    """
    columns = []
    if len(table.rows) == 0:
        return columns

    for cell in table.rows[0].cells:
        raw = normalize(cell.text)
        columns.append({
            "name": remove_year(raw),
            "year": extract_year(raw)
        })
    return columns


def extract_okveds(table):
    """
    Извлекает значения ОКВЕД из первого столбца (начиная со 2-й строки, т.к. 1-я - заголовок).
    """
    okveds = []
    if len(table.rows) <= 1:
        return okveds

    for row in table.rows[1:]:
        if len(row.cells) > 0:
            val = row.cells[0].text.strip()
            if val:
                okveds.append(val)
    return okveds


def build_schema(blocks):
    """
    Строит схему данных на основе ПЕРВОЙ таблицы каждого блока.
    Предполагается, что структура колонок одинакова для всех таблиц в блоке.
    """
    schema = {}

    for block in blocks:
        if not block["tables"]:
            print(f"[WARN] Блок '{block['name']}' не содержит таблиц. Пропускаем.")
            continue

        # Берем первую таблицу для определения структуры колонок
        first_table = block["tables"][0]
        
        schema[block["name"]] = {
            "columns": extract_columns(first_table),
            "okveds": extract_okveds(first_table)
        }
        
        # Для отладки можно вывести количество колонок
        # print(f"  -> Columns detected: {len(schema[block['name']]['columns'])}")

    return schema


# --- ГЕНЕРАЦИЯ ТЕГОВ И ЗАПОЛНЕНИЕ ---

def make_tag(okved, column):
    """Создает тег вида [ОКВЕД][Имя_колонки][Год] или [ОКВЕД][Имя_колонки]."""
    if column["year"]:
        return f"[{okved}][{column['name']}][{column['year']}]"
    return f"[{okved}][{column['name']}]"


def fill_blocks(blocks, schema):
    """
    Проходит по всем таблицам во всех блоках и заменяет ячейки на теги.
    """
    for block in blocks:
        block_name = block["name"]
        
        if block_name not in schema:
            continue
            
        cols = schema[block_name]["columns"]
        num_cols = len(cols)

        for table in block["tables"]:
            # Проходим по строкам, начиная со второй (индекс 1), т.к. 0 - заголовок
            for row in table.rows[1:]:
                if len(row.cells) == 0:
                    continue
                    
                okved = row.cells[0].text.strip()
                if not okved:
                    continue # Пропускаем строки без ОКВЕД

                # Проходим по ячейкам строки
                for j, cell in enumerate(row.cells):
                    # Пропускаем первый столбец (ОКВЕД) и выходящие за границы схемы
                    if j == 0 or j >= num_cols:
                        continue

                    tag = make_tag(okved, cols[j])
                    cell.text = tag


# --- ГЛАВНЫЙ ЦИКЛ (MAIN) ---

def main():
    input_file = "input.docx"
    mapping_file = "mapping.csv"
    output_file = "output.docx"

    print(f"Loading mapping from {mapping_file}...")
    mapping_names = load_mapping(mapping_file)
    
    if not mapping_names:
        print("[ERROR] Список имен таблиц пуст. Проверьте mapping.csv")
        return

    print(f"Loaded {len(mapping_names)} table names.")
    
    print(f"Opening document {input_file}...")
    try:
        doc = Document(input_file)
    except FileNotFoundError:
        print(f"[ERROR] Файл {input_file} не найден.")
        return

    print("Detecting blocks...")
    blocks = detect_blocks(doc, mapping_names)
    print(f"Found {len(blocks)} blocks.")

    if not blocks:
        print("[WARN] Блоки не найдены. Проверьте заголовки в документе и mapping.csv")
        # Можно сохранить файл как есть или выйти
        # doc.save(output_file)
        return

    print("Building schema...")
    schema = build_schema(blocks)

    print("Filling blocks with tags...")
    fill_blocks(blocks, schema)

    print(f"Saving result to {output_file}...")
    doc.save(output_file)
    print("Done!")

    # Отчет для пользователя
    print("\n--- SUMMARY ---")
    for b in blocks:
        count = len(b["tables"])
        print(f"Block: {b['name']} -> Tables: {count}")


if __name__ == "__main__":
    main()
