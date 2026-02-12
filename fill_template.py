# -*- coding: utf-8 -*-
"""
Заполнение шаблона «Шаблон Район нагрузок электрической сети.docx»
только по ячейкам из справочника индексов (test.ipynb).
Поддержка расположения подстанций: координаты (row, col) в doc.tables[0], отображение кружком (PNG).
"""

import os
import json
import zlib
import struct
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.shared import Inches, Pt

try:
    from PIL import Image, ImageDraw, ImageFont
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False


FONT_NAME = "Times New Roman"


def set_cell_text(cell, value):
    """Записать текст в ячейку таблицы (очистить параграф и вставить значение). Шрифт: Times New Roman."""
    if value is None:
        value = ""
    value = str(value).strip()
    if cell.paragraphs:
        p = cell.paragraphs[0]
        p.clear()
        run = p.add_run(value)
        run.font.name = FONT_NAME
    else:
        p = cell.add_paragraph(value)
        if p.runs:
            p.runs[0].font.name = FONT_NAME


# --- Справочник индексов (таблица, строка, столбец) для каждого поля ---
# Структура документа: 0 — План, 1 — Масштаб, 2 — Нагрузки, 3 — Напряжения, 4 — Доп данные, 5 — Условия.
LOADS_TABLE_INDEX = 2
VOLTAGE_TABLE_INDEX = 3
ADDITIONAL_TABLE_INDEX = 4
CONDITIONS_TABLE_INDEX = 5

# Таблица нагрузок (индекс 2): данные о нагрузках района (строки 4–7 — подстанции, строка 8 — генерация).
# Для каждой подстанции заполняется только один уровень напряжения: либо 110 кВ (колонки 1–4), либо 10 кВ (колонки 5–8).
TABLE1_LOAD_ROWS = list(range(4, 8))   # 4, 5, 6, 7
TABLE1_LOAD_COLS = {
    "number_ps": 0,           # номер подстанции на плане
    "p_max_110": 1,           # P макс 110 кВ, МВт
    "tg_max_110": 2,          # tgφ макс 110 кВ
    "p_min_110": 3,           # P мин 110 кВ, МВт
    "tg_min_110": 4,          # tgφ мин 110 кВ
    "p_max_10": 5,            # P макс 10 кВ, МВт
    "tg_max_10": 6,           # tgφ макс 10 кВ
    "p_min_10": 7,            # P мин 10 кВ, МВт
    "tg_min_10": 8,           # tgφ мин 10 кВ
    "reliability": 9,         # состав нагрузки по категориям надёжности
    "t_ma": 10,               # Тма, ч
}
# Колонки по уровням напряжения: при заполнении одного уровня второй очищаем
TABLE1_COLS_110 = (1, 2, 3, 4)   # p_max_110, tg_max_110, p_min_110, tg_min_110
TABLE1_COLS_10 = (5, 6, 7, 8)    # p_max_10, tg_max_10, p_min_10, tg_min_10
TABLE1_GEN_ROW = 8
TABLE1_GEN_CELLS = {
    "gen_node": 4,            # узел с генерацией
    "p_ust": 7,               # Руст, МВт
    "gen_count": 10,          # количество генераторов
}

# Таблица 2: уровни напряжения на шинах ПС А
TABLE2_PLACES = [
    (0, 1, "u_max"),          # режим максимальных нагрузок
    (1, 1, "u_min"),          # режим минимальных нагрузок
]

# Таблица 3: дополнительные данные 3.4.1–3.4.5
TABLE3_PLACES = [
    (0, 10, "km"),            # Км
    (1, 7, "h"),              # h, ч
    (2, 2, "t_r"),            # ТР
    (3, 1, "e_n"),            # ЕН
    (4, 2, "ts_e"),           # ЦЭ
]

# Таблица 4: условия сооружения и работы сети 3.4.6
TABLE4_PLACES = [
    (0, 2, "oes"),            # район принадлежит ОЭС
    (1, 2, "theta_ohl"),      # ϑохл
]

# Размер кружка в дюймах при вставке в ячейку (по умолчанию; можно задать в config/data как circle_size_inches)
CIRCLE_SIZE_INCHES = 0.18


def _make_png_chunk(chunk_type, data):
    """Собрать PNG-чанк: длина (4), тип (4), данные, CRC32 (4)."""
    chunk = chunk_type + data
    crc = zlib.crc32(chunk) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + chunk + struct.pack(">I", crc)


def _make_circle_png(size_px=200, line_width=10):
    """
    Создать PNG: прозрачный фон, только чёрный контур кружка (без заливки внутри).
    size_px — сторона квадрата, line_width — толщина контура (пиксели).
    """
    w = h = size_px
    cx, cy = w / 2.0, h / 2.0
    r_outer = min(cx, cy) - 1
    r_inner = max(0, r_outer - line_width)
    rows = []
    for y in range(h):
        row = bytearray([0])  # filter byte
        for x in range(w):
            d = (x - cx) ** 2 + (y - cy) ** 2
            if d <= r_inner ** 2:
                row.extend([0, 0, 0, 0])   # внутри — прозрачно
            elif d <= r_outer ** 2:
                row.extend([0, 0, 0, 255])  # только контур — чёрный
            else:
                row.extend([0, 0, 0, 0])    # снаружи — прозрачно
        rows.append(bytes(row))
    raw = b"".join(rows)
    compressed = zlib.compress(raw, 9)
    ihdr = struct.pack(">2I5B", w, h, 8, 6, 0, 0, 0)  # 8 bit, RGBA, no filter, no interlace
    signature = b"\x89PNG\r\n\x1a\n"
    ihdr_chunk = _make_png_chunk(b"IHDR", ihdr)
    idat_chunk = _make_png_chunk(b"IDAT", compressed)
    iend_chunk = _make_png_chunk(b"IEND", b"")
    return signature + ihdr_chunk + idat_chunk + iend_chunk


def _draw_label_center_png(png_bytes, label):
    """Нарисовать подпись в центре PNG (кружка). Требует Pillow. Возвращает PNG-байты."""
    if not _PIL_AVAILABLE or not label:
        return png_bytes
    label = str(label).strip()
    img = Image.open(BytesIO(png_bytes)).convert("RGBA")
    w, h = img.size
    draw = ImageDraw.Draw(img)
    font_size = max(20, min(w, h) // 3)
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except (OSError, IOError):
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", font_size)
        except (OSError, IOError):
            font = ImageFont.load_default()
    if hasattr(draw, "textbbox"):
        bbox = draw.textbbox((0, 0), label, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    else:
        tw, th = draw.textsize(label, font=font)
    x = (w - tw) // 2
    y = (h - th) // 2
    draw.text((x, y), label, fill=(0, 0, 0, 255), font=font)
    out = BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def add_circle_to_cell(cell, label=None, size_inches=None):
    """Вставить в ячейку кружок (PNG); подпись — в центре кружка (если есть Pillow)."""
    size_inches = size_inches or CIRCLE_SIZE_INCHES
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    png_bytes = _make_circle_png(100, 5)
    if label:
        png_bytes = _draw_label_center_png(png_bytes, label)
    run.add_picture(
        BytesIO(png_bytes),
        width=Inches(size_inches),
        height=Inches(size_inches),
    )


def fill_table0_substations(doc, positions, draw_circles=False, circle_size_inches=None):
    """
    Наложить кружки подстанций на таблицу doc.tables[0] по координатам (row, col).
    positions: список dict с ключами row, col, label (индексы ячейки и подпись: «ПС А», «1», «2», …).
    Подпись рисуется в центре кружка (при наличии Pillow).
    """
    if not positions or len(doc.tables) < 1 or not draw_circles:
        return
    size = circle_size_inches if circle_size_inches is not None else CIRCLE_SIZE_INCHES
    tbl = doc.tables[0]
    for pos in positions:
        row_idx = pos.get("row", 0)
        col_idx = pos.get("col", 0)
        label = pos.get("label", "")
        if row_idx < 0 or row_idx >= len(tbl.rows):
            continue
        row = tbl.rows[row_idx]
        if col_idx < 0 or col_idx >= len(row.cells):
            continue
        cell = row.cells[col_idx]
        add_circle_to_cell(cell, label=label, size_inches=size)


def fill_table1(doc, data):
    """Таблица нагрузок района и генерация (doc.tables[2])."""
    idx = LOADS_TABLE_INDEX
    if len(doc.tables) <= idx:
        return
    t = doc.tables[idx]
    load = data.get("loads") or []
    has_110 = ("p_max_110", "tg_max_110", "p_min_110", "tg_min_110")
    has_10 = ("p_max_10", "tg_max_10", "p_min_10", "tg_min_10")
    for i, row_idx in enumerate(TABLE1_LOAD_ROWS):
        if i >= len(load):
            break
        row_data = load[i]
        row_has_110 = any(k in row_data for k in has_110)
        row_has_10 = any(k in row_data for k in has_10)
        for key, col in TABLE1_LOAD_COLS.items():
            if key in row_data and row_idx < len(t.rows) and col < len(t.rows[row_idx].cells):
                set_cell_text(t.rows[row_idx].cells[col], row_data[key])
        # Очистить только те ячейки, для которых в этой строке нет данных
        if row_idx < len(t.rows):
            row_cells = t.rows[row_idx].cells
            if not row_has_110:
                for col in TABLE1_COLS_110:
                    if col < len(row_cells):
                        set_cell_text(row_cells[col], "")
            if not row_has_10:
                for col in TABLE1_COLS_10:
                    if col < len(row_cells):
                        set_cell_text(row_cells[col], "")

    gen = data.get("generation") or {}
    if TABLE1_GEN_ROW < len(t.rows):
        row = t.rows[TABLE1_GEN_ROW]
        for key, col in TABLE1_GEN_CELLS.items():
            if key in gen and col < len(row.cells):
                set_cell_text(row.cells[col], gen[key])


def fill_table2(doc, data):
    """Таблица напряжений на шинах ПС А (doc.tables[3])."""
    idx = VOLTAGE_TABLE_INDEX
    if len(doc.tables) <= idx:
        return
    t = doc.tables[idx]
    for row_idx, col_idx, key in TABLE2_PLACES:
        if key not in data:
            continue
        if row_idx < len(t.rows) and col_idx < len(t.rows[row_idx].cells):
            set_cell_text(t.rows[row_idx].cells[col_idx], data[key])


def fill_table3(doc, data):
    """Таблица дополнительных данных 3.4 (doc.tables[4])."""
    idx = ADDITIONAL_TABLE_INDEX
    if len(doc.tables) <= idx:
        return
    t = doc.tables[idx]
    for row_idx, col_idx, key in TABLE3_PLACES:
        if key not in data:
            continue
        if row_idx < len(t.rows) and col_idx < len(t.rows[row_idx].cells):
            set_cell_text(t.rows[row_idx].cells[col_idx], data[key])


def fill_table4(doc, data):
    """Таблица условий сооружения и работы сети (doc.tables[5])."""
    idx = CONDITIONS_TABLE_INDEX
    if len(doc.tables) <= idx:
        return
    t = doc.tables[idx]
    for row_idx, col_idx, key in TABLE4_PLACES:
        if key not in data:
            continue
        if row_idx < len(t.rows) and col_idx < len(t.rows[row_idx].cells):
            set_cell_text(t.rows[row_idx].cells[col_idx], data[key])


def set_paragraph3_text(doc, value):
    """
    Заполнить doc.paragraphs[3]: сохранить текст без последних 4 символов и дописать value.
    Шрифт Times New Roman; подчёркнут только вставляемый текст (value).
    """
    if len(doc.paragraphs) <= 3:
        return
    p = doc.paragraphs[3]
    base_text = (p.text or "")[:-4]
    insert_text = str(value).strip()
    p.clear()
    if base_text:
        r1 = p.add_run(base_text)
        r1.font.name = FONT_NAME
    if insert_text:
        r2 = p.add_run(insert_text)
        r2.font.name = FONT_NAME
        r2.font.underline = WD_UNDERLINE.SINGLE


def fill_paragraph3(doc, data):
    """Заполнить параграф 3 (на курсовой проект № …) из data['paragraph3'] или data['course_project_number']."""
    value = data.get("paragraph3") or data.get("course_project_number")
    if value is None:
        return
    set_paragraph3_text(doc, value)


def set_paragraph4_text(doc, value1, value2):
    """
    Заполнить doc.paragraphs[4]: текст по шаблону text[0:8] + ' value1 ' + text[16:22] + ' value2 ' + text[29:39].
    Перед и после вставляемых элементов — пробелы; только вставляемые элементы подчёркнуты. Шрифт Times New Roman.
    """
    if len(doc.paragraphs) <= 4:
        return
    p = doc.paragraphs[4]
    t = p.text or ""
    v1 = str(value1).strip()
    v2 = str(value2).strip()
    p.clear()

    def add_run(text, underline=False):
        if not text:
            return
        r = p.add_run(text)
        r.font.name = FONT_NAME
        if underline:
            r.font.underline = WD_UNDERLINE.SINGLE

    add_run(t[0:8])
    add_run(" " + v1 + " ", underline=True)
    add_run(t[16:22])
    add_run(" " + v2 + " ", underline=True)
    add_run(t[29:39])


def fill_paragraph4(doc, data):
    """Заполнить параграф 4 из data['paragraph4_1'] и data['paragraph4_2'] (или paragraph4_val1, paragraph4_val2)."""
    val1 = data.get("paragraph4_1") or data.get("paragraph4_val1")
    val2 = data.get("paragraph4_2") or data.get("paragraph4_val2")
    if val1 is None or val2 is None:
        return
    set_paragraph4_text(doc, val1, val2)


def set_paragraph5_text(doc, value):
    """
    Заполнить doc.paragraphs[5] текстом value. Перед и после — по «__», весь фрагмент подчёркнут. Шрифт Times New Roman.
    """
    if len(doc.paragraphs) <= 5:
        return
    p = doc.paragraphs[5]
    text = str(value).strip() if value is not None else ""
    p.clear()
    full_text = "__" + text + "__"
    if full_text.strip():
        r = p.add_run(full_text)
        r.font.name = FONT_NAME
        r.font.underline = WD_UNDERLINE.SINGLE


def fill_paragraph5(doc, data):
    """Заполнить параграф 5 (ФИО/имя) из data['paragraph5'] или data['name']."""
    value = data.get("paragraph5") or data.get("name")
    if value is None:
        return
    set_paragraph5_text(doc, value)


def fill_scale(doc, data):
    """Записать масштаб плана в ячейку, если заданы scale_text и scale_cell [table_idx, row, col]."""
    scale_text = data.get("scale_text")
    scale_cell = data.get("scale_cell")
    if not scale_text or not scale_cell or not isinstance(scale_cell, (list, tuple)) or len(scale_cell) != 3:
        return
    ti, ri, ci = int(scale_cell[0]), int(scale_cell[1]), int(scale_cell[2])
    if ti < 0 or ti >= len(doc.tables):
        return
    t = doc.tables[ti]
    if ri < 0 or ri >= len(t.rows) or ci < 0 or ci >= len(t.rows[ri].cells):
        return
    set_cell_text(t.rows[ri].cells[ci], scale_text)


def fill_document(doc, data):
    """Заполнить все таблицы по справочнику индексов. Кружки на таблицу 0 — только при draw_substation_circles: true."""
    fill_paragraph3(doc, data)
    fill_paragraph4(doc, data)
    fill_paragraph5(doc, data)
    fill_table0_substations(
        doc,
        data.get("substation_positions") or [],
        draw_circles=data.get("draw_substation_circles", False),
        circle_size_inches=data.get("circle_size_inches"),
    )
    fill_scale(doc, data)
    fill_table1(doc, data)
    fill_table2(doc, data.get("voltage_ps_a") or {})
    fill_table3(doc, data.get("additional") or {})
    fill_table4(doc, data.get("conditions") or {})


def load_data(path):
    """Загрузить данные из JSON (UTF-8)."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    import argparse
    script_dir = os.path.dirname(os.path.abspath(__file__))

    parser = argparse.ArgumentParser(description="Заполнение шаблона по справочнику индексов")
    parser.add_argument("--template", "-t", default=None,
                        help="Путь к шаблону .docx")
    parser.add_argument("--data", "-d", default=None,
                        help="Путь к JSON с данными для подстановки")
    parser.add_argument("--out", "-o", default=None,
                        help="Путь к выходному файлу")
    args = parser.parse_args()

    template_path = args.template or os.path.join(script_dir, "Шаблон Район нагрузок электрической сети.docx")
    data_path = args.data or os.path.join(script_dir, "data_template.json")
    out_path = args.out or os.path.join(script_dir, "result.docx")

    if not os.path.isfile(template_path):
        print(f"Шаблон не найден: {template_path}")
        return 1

    if not os.path.isfile(data_path):
        print(f"Файл данных не найден: {data_path}")
        print("Создан пример data_template.json — отредактируйте его и запустите снова.")
        create_example_data(data_path)
        return 1

    data = load_data(data_path)
    doc = Document(template_path)
    fill_document(doc, data)
    doc.save(out_path)
    print(f"Сохранено: {out_path}")
    return 0


def create_example_data(path):
    """Создать пример JSON с описанием полей."""
    example = {
        "loads": [
            {
                "number_ps": "ПС-1",
                "p_max_110": "25",
                "tg_max_110": "0.33",
                "p_min_110": "12",
                "tg_min_110": "0.35",
                "p_max_10": "20",
                "tg_max_10": "0.33",
                "p_min_10": "10",
                "tg_min_10": "0.35",
                "reliability": "I – 80%, II – 15%, III – 5%",
                "t_ma": "5000"
            },
            {
                "number_ps": "ПС-2",
                "p_max_110": "30",
                "tg_max_110": "0.33",
                "p_min_110": "15",
                "tg_min_110": "0.35",
                "p_max_10": "24",
                "tg_max_10": "0.33",
                "p_min_10": "12",
                "tg_min_10": "0.35",
                "reliability": "I – 70%, II – 25%, III – 5%",
                "t_ma": "5200"
            },
            {
                "number_ps": "ПС-3",
                "p_max_110": "20",
                "tg_max_110": "0.33",
                "p_min_110": "10",
                "tg_min_110": "0.35",
                "p_max_10": "16",
                "tg_max_10": "0.33",
                "p_min_10": "8",
                "tg_min_10": "0.35",
                "reliability": "I – 90%, II – 10%",
                "t_ma": "4800"
            },
            {
                "number_ps": "ПС-4",
                "p_max_110": "18",
                "tg_max_110": "0.33",
                "p_min_110": "9",
                "tg_min_110": "0.35",
                "p_max_10": "14",
                "tg_max_10": "0.33",
                "p_min_10": "7",
                "tg_min_10": "0.35",
                "reliability": "I – 85%, II – 15%",
                "t_ma": "5100"
            }
        ],
        "generation": {
            "gen_node": "ПС-2",
            "p_ust": "12",
            "gen_count": "2"
        },
        "voltage_ps_a": {
            "u_max": "116",
            "u_min": "112"
        },
        "additional": {
            "km": "0.95",
            "h": "4",
            "t_r": "25",
            "e_n": "0.1",
            "ts_e": "2.5"
        },
        "conditions": {
            "oes": "Центр",
            "theta_ohl": "-25"
        },
        "substation_positions": [
            {"row": 2, "col": 2, "label": "ПС А"},
            {"row": 2, "col": 6, "label": "1"},
            {"row": 5, "col": 3, "label": "2"},
            {"row": 5, "col": 8, "label": "3"},
            {"row": 8, "col": 5, "label": "4"}
        ]
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(example, f, ensure_ascii=False, indent=2)
    print(f"Создан файл: {path}")


if __name__ == "__main__":
    exit(main() or 0)
