#!/usr/bin/env python3
"""
Динамическое сканирование и построение маппинга соответствия
для каждой ячейки Word таблиц с Excel источниками
"""

import pandas as pd
from docx import Document
from pathlib import Path
import json
from typing import Dict, List, Tuple, Optional
import re

class DynamicMappingBuilder:
    """Класс для динамического построения маппинга Word таблиц с Excel источниками"""

    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.input_dir = base_dir / "input"
        self.mappings_dir = self.input_dir / "mappings"
        self.excel_dir = self.input_dir / "excel"
        self.output_dir = base_dir / "output"

        # Загружаем справочники
        self.table_sources = self._load_table_sources()
        self.column_mappings = self._load_column_mappings()
        self.okved_map = self._load_okved_map()

    def _load_table_sources(self) -> Dict[str, str]:
        """Загружаем маппинг таблиц к Excel файлам"""
        file_path = self.mappings_dir / "table_source_data_mapping.csv"
        df = pd.read_csv(file_path, sep=';')
        mapping = {}
        for _, row in df.iterrows():
            table_name = row['Таблица'].strip('"')
            excel_file = row['Источник']
            mapping[table_name] = excel_file
        print(f"📂 Загружены источники таблиц: {mapping}")
        return mapping

    def _load_column_mappings(self) -> Dict[str, List[Dict]]:
        """Загружаем маппинг показателей к колонкам Excel"""
        file_path = self.mappings_dir / "column_mapping_v2.csv"
        df = pd.read_csv(file_path, sep=';')

        mapping = {}
        for _, row in df.iterrows():
            excel_file = row['Excel файл']
            if excel_file not in mapping:
                mapping[excel_file] = []

            indicator_data = {
                'indicator_name': row['Название показателя'],
                'indicator_code': row['Код показателя'],
                'col_2022': int(row['Excel колонка 2022']) if pd.notna(row['Excel колонка 2022']) else None,
                'col_2023': int(row['Excel колонка 2023']) if pd.notna(row['Excel колонка 2023']) else None,
                'keyword_2022': row['Ключевое слово 2022'],
                'keyword_2023': row['Ключевое слово 2023']
            }
            mapping[excel_file].append(indicator_data)

        return mapping

    def _load_okved_map(self) -> Dict[str, str]:
        """Загружаем маппинг ОКВЭД кодов"""
        file_path = self.mappings_dir / "okved_mapping.csv"
        df = pd.read_csv(file_path, sep=';')
        return dict(zip(df['Код'], df['Наименование']))

    def scan_word_table(self, table, table_idx: int) -> Dict:
        """Сканирует Word таблицу и определяет структуру данных"""
        table_info = {
            'table_index': table_idx,
            'rows': len(table.rows),
            'cols': len(table.rows[0].cells) if table.rows else 0,
            'headers': [],
            'data_structure': []
        }

        # Анализируем заголовки (первые несколько строк)
        header_rows = min(5, len(table.rows))
        for row_idx in range(header_rows):
            row_data = []
            for col_idx in range(len(table.rows[row_idx].cells)):
                cell_text = table.rows[row_idx].cells[col_idx].text.strip()
                row_data.append(cell_text)
            table_info['headers'].append(row_data)

        # Определяем структуру данных
        table_info['data_structure'] = self._analyze_table_structure(table_info['headers'])
        
        print(f"  📋 Найденные показатели: {[ind['name'] for ind in table_info['data_structure']['indicators']][:5]}...")  # Первые 5 показателей

        return table_info

    def _get_table_context(self, doc, table_idx: int) -> str:
        """Получает контекст перед таблицей (текст ближайших параграфов)"""
        context_paragraphs = []

        # Собираем параграфы перед таблицей (последние 5-10 параграфов)
        paragraphs = doc.paragraphs
        start_idx = max(0, table_idx * 2 - 10)  # Примерная позиция
        end_idx = table_idx * 2  # До позиции таблицы

        for i in range(start_idx, min(end_idx, len(paragraphs))):
            text = paragraphs[i].text.strip()
            if text and len(text) > 5:  # Игнорируем пустые и очень короткие параграфы
                context_paragraphs.append(text)

        # Если не нашли достаточно контекста, пробуем другой подход
        if len(context_paragraphs) < 2:
            # Ищем параграфы, содержащие номера таблиц
            for para in paragraphs:
                text = para.text.strip()
                if text and any(f"{i}." in text for i in range(1, 20)):
                    context_paragraphs.append(text)
                    if len(context_paragraphs) >= 5:
                        break

        return " ".join(context_paragraphs[-5:])  # Возвращаем последние 5 параграфов

    def _analyze_table_structure(self, headers: List[List[str]]) -> Dict:
        """Анализирует структуру таблицы на основе заголовков"""
        structure = {
            'has_okved_column': False,
            'okved_column_idx': None,
            'indicators': [],
            'years': [],
            'data_start_row': 0
        }

        # Ищем колонку с ОКВЭД (обычно это первая колонка или колонка с "Код")
        for row_idx, row in enumerate(headers):
            for col_idx, cell in enumerate(row):
                if 'Код' in cell or 'ОКВЭД' in cell or (col_idx == 0 and len(cell) > 0):
                    structure['has_okved_column'] = True
                    structure['okved_column_idx'] = col_idx
                    structure['data_start_row'] = max(row_idx + 1, 5)  # Начинаем после первых 5 строк
                    break

        # Если так и не нашли - сначала первая колонка
        if not structure['has_okved_column']:
            structure['has_okved_column'] = True
            structure['okved_column_idx'] = 0
            structure['data_start_row'] = 5

        # Ищем показатели и года в заголовках
        for row_idx, row in enumerate(headers):
            for col_idx, cell in enumerate(row):
                cell_lower = cell.lower()

                # Пропускаем первую колонку (ОКВЭД) и пустые ячейки
                if col_idx == structure['okved_column_idx'] or not cell:
                    continue

                # Ищем года
                if '2022' in cell or '2023' in cell or 'предыдущий' in cell or 'отчетный' in cell:
                    if '2022' in cell or 'предыдущий' in cell:
                        structure['years'].append({'year': 2022, 'col': col_idx, 'row': row_idx})
                    elif '2023' in cell or 'отчетный' in cell:
                        structure['years'].append({'year': 2023, 'col': col_idx, 'row': row_idx})
                # Ищем показатели (если это не год и не пусто, и не число)
                elif cell and not any(c.isdigit() for c in cell[:3]):  # Не начинается с цифр
                    # Пропускаем очень короткие или похожие на заголовки текст
                    if len(cell) > 2 and cell not in ['на конец года, тысяч рублей', 'Продолжение таблицы']:
                        # Проверяем, что это не повторение предыдущего показателя
                        if not structure['indicators'] or structure['indicators'][-1]['name'] != cell:
                            structure['indicators'].append({
                                'name': cell,
                                'col': col_idx,
                                'row': row_idx
                            })

        print(f"    📊 Структура: {len(structure['indicators'])} показателей, {len(structure['years'])} лет, данные с ряда {structure['data_start_row']}")
        
        return structure

    def find_excel_source_for_table(self, table_context: str, table_info: Dict) -> Optional[str]:
        """Находит Excel источник для Word таблицы по номеру таблицы"""
        table_idx = table_info['table_index']

        # Маппинг индекса таблицы к номеру источника
        # Таблица 0 -> 1, Таблица 1 -> 3, Таблица 2 -> 4, etc.
        table_number_mapping = {
            0: "1", 1: "3", 2: "4", 3: "5", 4: "6", 5: "7", 6: "8"
        }

        if table_idx in table_number_mapping:
            table_number = table_number_mapping[table_idx]
            source_key = f"{table_number}. "

            # Ищем источник, который начинается с этого номера
            for source_title, excel_file in self.table_sources.items():
                if source_title.startswith(source_key):
                    print(f"  🎯 Найден источник по номеру таблицы: {excel_file}")
                    return excel_file

        print(f"  ❌ Нет источника для таблицы {table_idx + 1}")
        return None

    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """Вычисляет схожесть двух текстов"""
        text1_words = set(text1.lower().split())
        text2_words = set(text2.lower().split())

        intersection = text1_words.intersection(text2_words)
        union = text1_words.union(text2_words)

        return len(intersection) / len(union) if union else 0

    def build_cell_mapping(self, table_info: Dict, excel_source: str) -> Dict:
        """Строит маппинг для каждой ячейки таблицы"""
        mapping = {
            'table_index': table_info['table_index'],
            'excel_source': excel_source,
            'cell_mappings': []
        }

        structure = table_info['data_structure']

        # Гарантируем, что есть ОКВЭД колонка (первая колонка)
        if not structure['has_okved_column']:
            structure['has_okved_column'] = True
            structure['okved_column_idx'] = 0
            structure['data_start_row'] = max(5, len(structure['indicators']))  # Начинаем после заголовков

        # Для каждой строки данных
        for row_idx in range(structure['data_start_row'], table_info['rows']):
            okved_code = None

            # Получаем ОКВЭД код из первой колонки (обычно это номер ОКВЭД)
            try:
                # Находим первую непустую ячейку в первой колонке после ОКВЭД заголовка
                for search_row in range(row_idx, min(row_idx + 1, table_info['rows'])):
                    if structure['okved_column_idx'] is not None and structure['okved_column_idx'] < len(table_info['headers'][0]):
                        okved_cell = table_info.get('raw_table_data', {}).get((search_row, structure['okved_column_idx']), '')
                        if okved_cell:
                            okved_code = okved_cell.strip()
                            break
            except (IndexError, KeyError):
                pass

            # Для каждого показателя
            for indicator in structure['indicators']:
                indicator_name = indicator['name']

                # Находим соответствующий маппинг в Excel
                excel_mapping = self._find_excel_indicator_mapping(excel_source, indicator_name)

                if excel_mapping:
                    # Для каждого года
                    for year_info in structure['years']:
                        cell_mapping = {
                            'row': row_idx,
                            'col': indicator['col'],
                            'okved_code': okved_code,
                            'indicator_name': indicator_name,
                            'indicator_code': excel_mapping.get('indicator_code', ''),
                            'year': year_info['year'],
                            'excel_col': excel_mapping.get(f'col_{year_info["year"]}', ''),
                            'confidence': 1.0  # Пока фиксированная уверенность
                        }
                        mapping['cell_mappings'].append(cell_mapping)

        return mapping

    def _find_excel_indicator_mapping(self, excel_file: str, indicator_name: str) -> Optional[Dict]:
        """Находит маппинг показателя в Excel файле"""
        if excel_file not in self.column_mappings:
            print(f"    ⚠️ Нет маппингов для файла {excel_file}")
            return None

        # Ищем точное совпадение
        for mapping in self.column_mappings[excel_file]:
            if mapping['indicator_name'].lower() == indicator_name.lower():
                print(f"    ✅ Найдено точное совпадение: {mapping['indicator_name']}")
                return mapping

        # Ищем fuzzy совпадение
        best_match = None
        best_score = 0

        for mapping in self.column_mappings[excel_file]:
            score = self._calculate_similarity(indicator_name, mapping['indicator_name'])
            if score > best_score and score > 0.5:  # Понизили порог с 0.7 до 0.5
                best_score = score
                best_match = mapping

        if best_match:
            print(f"    ✅ Найдено fuzzy совпадение: {best_match['indicator_name']} (score: {best_score:.2f})")
        else:
            print(f"    ❌ Не найдено совпадение для '{indicator_name}'")
        
        return best_match

    def build_complete_mapping(self, word_file: Path) -> Dict:
        """Строит полный маппинг для всего документа"""
        print(f"🔍 Анализируем документ: {word_file}")

        doc = Document(word_file)
        complete_mapping = {
            'document': str(word_file.name),
            'tables': [],
            'summary': {
                'total_tables': len(doc.tables),
                'mapped_tables': 0,
                'total_cells': 0,
                'mapped_cells': 0
            }
        }

        for table_idx, table in enumerate(doc.tables):
            print(f"📊 Обрабатываем таблицу {table_idx + 1}/{len(doc.tables)}")

            # Получаем контекст перед таблицей (текст параграфов)
            table_context = self._get_table_context(doc, table_idx)
            table_info = self.scan_word_table(table, table_idx)

            # Находим Excel источник по контексту
            excel_source = self.find_excel_source_for_table(table_context, table_info)

            if excel_source:
                print(f"  ✅ Найден источник: {excel_source}")

                # Строим маппинг ячеек
                cell_mapping = self.build_cell_mapping(table_info, excel_source)

                complete_mapping['tables'].append({
                    'table_info': table_info,
                    'table_context': table_context,
                    'excel_source': excel_source,
                    'cell_mapping': cell_mapping
                })

                complete_mapping['summary']['mapped_tables'] += 1
                complete_mapping['summary']['mapped_cells'] += len(cell_mapping['cell_mappings'])
            else:
                print(f"  ❌ Источник не найден")
                complete_mapping['tables'].append({
                    'table_info': table_info,
                    'table_context': table_context,
                    'excel_source': None,
                    'cell_mapping': None
                })

            complete_mapping['summary']['total_cells'] += (table_info['rows'] * table_info['cols'])

        # Рассчитываем процент смаппированных ячеек
        complete_mapping['summary']['mapped_percentage'] = (
            complete_mapping['summary']['mapped_cells'] / complete_mapping['summary']['total_cells'] * 100
        ) if complete_mapping['summary']['total_cells'] > 0 else 0

        return complete_mapping

    def save_mapping(self, mapping: Dict, output_file: Path):
        """Сохраняет маппинг в JSON файл"""
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)

        print(f"💾 Маппинг сохранен в: {output_file}")

def main():
    """Основная функция"""
    base_dir = Path(__file__).parent.resolve()

    # Создаем билдер маппинга
    builder = DynamicMappingBuilder(base_dir)

    # Анализируем эталонный документ
    reference_file = base_dir / "output" / "как должно быть!.docx"
    if reference_file.exists():
        print("🎯 Строим динамический маппинг для эталонного документа...")
        mapping = builder.build_complete_mapping(reference_file)

        # Сохраняем маппинг
        output_file = base_dir / "output" / "dynamic_mapping_control.json"
        builder.save_mapping(mapping, output_file)

        # Выводим статистику
        summary = mapping['summary']
        print("\n📊 СТАТИСТИКА МАППИНГА:")
        print(f"  Всего таблиц: {summary['total_tables']}")
        print(f"  Смаппированных таблиц: {summary['mapped_tables']}")
        print(f"  Всего ячеек: {summary['total_cells']}")
        print(f"  Смаппированных ячеек: {summary['mapped_cells']}")
        print(f"  Процент смаппированных ячеек: {summary['mapped_percentage']:.1f}%")
    else:
        print(f"❌ Эталонный файл не найден: {reference_file}")

if __name__ == "__main__":
    main()