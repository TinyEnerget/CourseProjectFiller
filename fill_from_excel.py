# -*- coding: utf-8 -*-
"""
Заполнение документов по данным из Excel.
Столбец N → paragraph3, столбец "group" → paragraph4_1, "course" → paragraph4_2, "name" → paragraph5.
Остальные данные (нагрузки, таблицы) берутся из одного сгенерированного варианта.

Использование:
  python fill_from_excel.py --excel students.xlsx --template "Шаблон Район нагрузок электрической сети.docx" --out-dir variants
  python fill_from_excel.py --excel students.xlsx --config config_variants.json
"""

import os
import argparse

try:
    import openpyxl
except ImportError:
    openpyxl = None

from docx import Document

from fill_template import fill_document
from generate_variants import load_config, generate_one_variant


def load_rows_from_excel(path, sheet_name=None):
    """
    Загрузить строки из Excel. Первая строка — заголовки.
    Возвращает (headers_dict, rows), где headers_dict: имя столбца -> индекс (0-based),
    rows — список кортежей (значения ячеек строки).
    Столбец N — по заголовку «N» в первой строке.
    """
    if openpyxl is None:
        raise RuntimeError("Установите openpyxl: pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active if sheet_name is None else wb[sheet_name]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not all_rows:
        return {}, []
    header_row = all_rows[0]
    headers = {}
    for i, val in enumerate(header_row):
        if val is not None and str(val).strip():
            headers[str(val).strip()] = i
    rows = all_rows[1:]
    return headers, rows


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parser = argparse.ArgumentParser(
        description="Заполнение документов из Excel (столбец N → paragraph3, group → paragraph4_1, course → paragraph4_2, name → paragraph5)"
    )
    parser.add_argument("--excel", "-e", default=None,
                        help="Путь к Excel-файлу (по умолчанию students.xlsx)")
    parser.add_argument("--template", "-t", default=None,
                        help="Путь к шаблону .docx")
    parser.add_argument("--out-dir", "-d", default=None,
                        help="Папка для сохранения документов (по умолчанию — папка скрипта)")
    parser.add_argument("--config", "-c", default=None,
                        help="Путь к config_variants.json для генерации варианта")
    parser.add_argument("--sheet", "-s", default=None,
                        help="Имя листа в Excel (по умолчанию — первый)")
    parser.add_argument("--seed", type=int, default=None,
                        help="Seed для ГПСЧ при генерации варианта")
    args = parser.parse_args()

    excel_path = args.excel or os.path.join(script_dir, "students.xlsx")
    if not os.path.isfile(excel_path):
        print(f"Файл не найден: {excel_path}")
        return 1

    if openpyxl is None:
        print("Установите openpyxl: pip install openpyxl")
        return 1

    template_path = args.template or os.path.join(script_dir, "Шаблон Район нагрузок электрической сети.docx")
    if not os.path.isfile(template_path):
        print(f"Шаблон не найден: {template_path}")
        return 1

    config_path = args.config or os.path.join(script_dir, "config_variants.json")
    config = load_config(config_path) if os.path.isfile(config_path) else {}

    out_dir = args.out_dir or script_dir
    os.makedirs(out_dir, exist_ok=True)

    headers, rows = load_rows_from_excel(excel_path, sheet_name=args.sheet)
    def col_index(name):
        if headers.get(name) is not None:
            return headers[name]
        low = name.lower()
        for k, v in headers.items():
            if k and str(k).strip().lower() == low:
                return v
        return None
    idx_n = col_index("N")
    idx_group = col_index("group")
    idx_course = col_index("course")
    idx_name = col_index("name")

    if idx_n is None:
        print("Столбец с заголовком 'N' не найден в первой строке.")
    if idx_group is None:
        print("Столбец 'group' не найден в первой строке.")
    if idx_course is None:
        print("Столбец 'course' не найден в первой строке.")
    if idx_name is None:
        print("Столбец 'name' не найден в первой строке.")

    for row_num, row in enumerate(rows, start=2):
        if not row:
            continue
        def val(i):
            if i is None or i >= len(row):
                return None
            v = row[i]
            return None if v is None else str(v).strip() or None

        paragraph3 = val(idx_n)
        paragraph4_1 = val(idx_group)
        paragraph4_2 = val(idx_course)
        paragraph5 = val(idx_name)

        # Для каждой строки генерируем свой вариант задания (разный seed = разные нагрузки, ОЭС и т.д.)
        row_seed = (args.seed + row_num) if args.seed is not None else row_num
        variant_data = generate_one_variant(config, oes_override=None, seed=row_seed)
        data = dict(variant_data)
        if paragraph3 is not None:
            data["paragraph3"] = paragraph3
        if paragraph4_1 is not None:
            data["paragraph4_1"] = paragraph4_1
        if paragraph4_2 is not None:
            data["paragraph4_2"] = paragraph4_2
        if paragraph5 is not None:
            data["paragraph5"] = paragraph5
            data["name"] = paragraph5

        doc = Document(template_path)
        fill_document(doc, data)
        safe_name = (paragraph5 or str(row_num)).replace("/", "-").replace("\\", "-").strip() or f"row_{row_num}"
        out_name = f"{safe_name}.docx"
        out_path = os.path.join(out_dir, out_name)
        doc.save(out_path)
        print(f"  {out_name}  N={paragraph3}, group={paragraph4_1}, course={paragraph4_2}, name={paragraph5}")

    print(f"Создано документов: {len(rows)} в папке {out_dir}")
    return 0


if __name__ == "__main__":
    exit(main() or 0)
