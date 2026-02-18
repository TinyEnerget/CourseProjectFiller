# -*- coding: utf-8 -*-
"""
Генерация заданного количества вариантов заданий с различными случайными значениями
в заданных диапазонах. ОЭС задаётся списком с фиксированной зимней температурой ϑохл.

Использование: 
  python generate_variants.py --count 10
  python generate_variants.py --count 5 --oes "Сибирь"
  python generate_variants.py --count 20 --out-dir variants --config config_variants.json
"""

import os
import json
import random
import argparse
from docx import Document

from fill_template import fill_document, convert_docx_to_pdf


def load_config(path):
    """Загрузить конфиг вариантов (ОЭС и диапазоны)."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _rand_range(r):
    """Случайное значение в диапазоне [min, max]. r может быть числом — вернуть как есть."""
    if isinstance(r, (int, float)) and "min" not in str(r):
        return r
    if isinstance(r, dict) and "min" in r and "max" in r:
        lo, hi = r["min"], r["max"]
        if isinstance(lo, int) and isinstance(hi, int):
            return random.randint(lo, hi)
        return round(random.uniform(float(lo), float(hi)), 4)
    return r


def _format_number(x):
    """Форматировать число для подстановки в документ (убрать лишние нули)."""
    if isinstance(x, float):
        if x == int(x):
            return str(int(x))
        return str(round(x, 2)).rstrip("0").rstrip(".")
    return str(x)


# Типы станций для строки 8 таблицы нагрузок: Руст и количество блоков
STATION_TYPES = {
    "ГЭС": {
        "p_ust_range": (200, 500),      # МВт на блок
        "total_range": (1200, 3000),    # общая мощность МВт
        "blocks_range": (3, 6),
    },
    "АЭС": {
        "p_ust_range": (800, 1200),
        "total_range": (2400, 3600),
        "blocks_range": (3, 3),
    },
    "ГРЭС": {
        "p_ust_range": (100, 300),
        "total_range": (1200, 3600),
        "blocks_range": (3, 12),
    },
}


def _generate_station(config):
    """
    Случайный выбор типа станции (ГЭС, АЭС, ГРЭС) и расчёт Руст, количества блоков
    в заданных диапазонах. Узел с генерацией всегда 5.
    """
    st_name = random.choice(list(STATION_TYPES.keys()))
    st = STATION_TYPES[st_name]
    p_min, p_max = st["p_ust_range"]
    total_min, total_max = st["total_range"]
    b_min, b_max = st["blocks_range"]

    for _ in range(100):
        blocks = random.randint(b_min, b_max)
        # Подобрать Руст так, чтобы total = blocks * p_ust лежало в [total_min, total_max]
        # p_ust_min = total_min / blocks, p_ust_max = total_max / blocks, и в [p_min, p_max]
        pu_lo = max(p_min, total_min / blocks)
        pu_hi = min(p_max, total_max / blocks)
        if pu_lo > pu_hi:
            continue
        p_ust = random.uniform(pu_lo, pu_hi)
        total = blocks * p_ust
        if total_min <= total <= total_max and p_min <= p_ust <= p_max:
            return {
                "gen_node": "5",
                "p_ust": _format_number(round(p_ust, 0)),
                "gen_count": str(blocks),
                "station_type": st_name,
            }
    # Fallback
    p_ust = (total_min + total_max) / 2 / max(b_min, 1)
    blocks = max(b_min, min(b_max, int(total_max / p_ust)))
    return {
        "gen_node": "5",
        "p_ust": _format_number(round(p_ust, 0)),
        "gen_count": str(blocks),
        "station_type": st_name,
    }


def generate_one_variant(config, oes_override=None, seed=None):
    """
    Сгенерировать один вариант данных.
    oes_override — название ОЭС (из списка) или None — случайный выбор.
    """
    if seed is not None:
        random.seed(seed)

    oes_list = config.get("oes_list", [])
    if not oes_list:
        raise ValueError("В конфиге должен быть oes_list")
    names = [o["name"] for o in oes_list]
    if oes_override:
        ov = (oes_override or "").strip().lower()
        oes = next((o for o in oes_list if (o["name"] or "").strip().lower() == ov), None)
        if oes:
            pass
        else:
            oes = random.choice(oes_list)
    else:
        oes = random.choice(oes_list)

    ranges = config.get("ranges", {})
    load_r = ranges.get("load", {})
    lr = ranges.get("loads_per_variant", {"min": 4, "max": 4})
    n_loads = random.randint(lr["min"], lr["max"]) if isinstance(lr, dict) and "min" in lr and "max" in lr else 4
    n_loads = min(max(n_loads, 1), 4)

    reliability_templates = load_r.get("reliability_templates", ["I – II – III"])
    min_total_load = max(0, float(ranges.get("min_total_load", 200)))

    def _rand_min_less_than_max(max_val, range_spec, eps=0.01):
        """Случайное значение P мин так, чтобы P мин < P макс."""
        upper = max_val - eps
        if upper <= 0:
            return 0.0
        if not isinstance(range_spec, dict) or "min" not in range_spec or "max" not in range_spec:
            return round(min(max_val * 0.5, upper), 4)
        lo = float(range_spec["min"])
        hi = min(float(range_spec["max"]), upper)
        if hi <= lo:
            hi = upper
        val = round(random.uniform(lo, hi), 4)
        return min(val, upper)

    # Выбирается 1 из 4 узлов — в нём указывается нагрузка 110 кВ и 10 кВ; в остальных трёх — только 10 кВ.
    # Условия: у каждого узла P мин < P макс; сумма всех нагрузок, вставляемых в таблицу (все P макс по узлам), не менее min_total_load МВт.
    load_110_node_index = random.randint(0, n_loads - 1) if n_loads else 0
    loads_raw = []
    for i in range(n_loads):
        row = {
            "number_ps": f"{i + 1}",
            "reliability": random.choice(reliability_templates),
            "t_ma": _rand_range(load_r.get("t_ma", {"min": 4500, "max": 5500})),
        }
        # У всех узлов — нагрузка 10 кВ; P мин < P макс
        p_max_10 = _rand_range(load_r.get("p_max_10", {"min": 10, "max": 38}))
        row["p_max_10"] = p_max_10
        row["tg_max_10"] = _rand_range(load_r.get("tg_max_10", {"min": 0.28, "max": 0.42}))
        row["p_min_10"] = _rand_min_less_than_max(p_max_10, load_r.get("p_min_10", {"min": 5, "max": 18}))
        row["tg_min_10"] = _rand_range(load_r.get("tg_min_10", {"min": 0.30, "max": 0.40}))
        # Только у выбранного узла — дополнительно нагрузка 110 кВ; P мин < P макс
        if i == load_110_node_index:
            p_max_110 = _rand_range(load_r.get("p_max_110", {"min": 15, "max": 45}))
            row["p_max_110"] = p_max_110
            row["tg_max_110"] = _rand_range(load_r.get("tg_max_110", {"min": 0.28, "max": 0.42}))
            row["p_min_110"] = _rand_min_less_than_max(p_max_110, load_r.get("p_min_110", {"min": 6, "max": 22}))
            row["tg_min_110"] = _rand_range(load_r.get("tg_min_110", {"min": 0.30, "max": 0.40}))
        loads_raw.append(row)

    # Сумма всех нагрузок, вставляемых в таблицу: по каждому узлу P макс 10 кВ; у одного узла дополнительно P макс 110 кВ
    total_load = 0.0
    for row in loads_raw:
        total_load += float(row["p_max_10"])
        if "p_max_110" in row:
            total_load += float(row["p_max_110"])

    if total_load < min_total_load and total_load >= 0.01:
        # Масштабируем все мощности так, чтобы общая нагрузка стала не менее min_total_load
        scale = min_total_load / total_load
        for row in loads_raw:
            for key in ("p_max_10", "p_min_10", "p_max_110", "p_min_110"):
                if key in row and isinstance(row[key], (int, float)):
                    row[key] = round(row[key] * scale, 2)
            if row["p_min_10"] >= row["p_max_10"]:
                row["p_min_10"] = max(0, round(row["p_max_10"] * 0.4, 2))
            if "p_max_110" in row and row["p_min_110"] >= row["p_max_110"]:
                row["p_min_110"] = max(0, round(row["p_max_110"] * 0.4, 2))

    loads = [{k: _format_number(v) if isinstance(v, (int, float)) else v for k, v in row.items()} for row in loads_raw]

    # Генерация: узел всегда 5; Руст и количество блоков — по типу станции (ГЭС, АЭС, ГРЭС)
    generation = _generate_station(config)


    # Таблицы 2 и 3 — напряжения. Значение 1 (u_max) всегда строго больше значения 2 (u_min). Между таблицами — разные пары.
    v_r = ranges.get("voltage_ps_a", {})
    u_max_range = v_r.get("u_max", {"min": 114, "max": 118})
    u_min_range = v_r.get("u_min", {"min": 108, "max": 114})
    for _ in range(50):
        v_max_1 = _rand_range(u_max_range)
        v_min_1 = _rand_range(u_min_range)
        if isinstance(v_max_1, (int, float)) and isinstance(v_min_1, (int, float)) and v_max_1 > v_min_1:
            break
    else:
        v_max_1, v_min_1 = 116.0, 110.0
    u_max_1 = _format_number(v_max_1)
    u_min_1 = _format_number(v_min_1)
    voltage_ps_a = {"u_max": u_max_1, "u_min": u_min_1}
    for _ in range(50):
        v_max_2 = _rand_range(u_max_range)
        v_min_2 = _rand_range(u_min_range)
        if not (isinstance(v_max_2, (int, float)) and isinstance(v_min_2, (int, float)) and v_max_2 > v_min_2):
            continue
        u_max_2 = _format_number(v_max_2)
        u_min_2 = _format_number(v_min_2)
        if u_max_2 != u_max_1 and u_min_2 != u_min_1:
            voltage_table3 = {"u_max": u_max_2, "u_min": u_min_2}
            break
    else:
        # принудительно разные от таблицы 2 и строго u_max > u_min
        v1, v2 = float(u_max_1), float(u_min_1)
        delta = max(1, (v1 - v2) / 4)
        v_max_2 = v1 + random.choice([-1, 1]) * delta
        v_min_2 = v2 + random.choice([-1, 1]) * delta
        if v_max_2 <= v_min_2:
            v_min_2 = v_max_2 - delta
        v_max_2 = max(109, min(119, v_max_2))
        v_min_2 = max(107, min(115, v_min_2))
        if v_max_2 <= v_min_2:
            v_min_2 = v_max_2 - 1
        u_max_2 = _format_number(v_max_2)
        u_min_2 = _format_number(v_min_2)
        voltage_table3 = {"u_max": u_max_2, "u_min": u_min_2}

    add_r = ranges.get("additional", {})
    additional = {
        "km": _format_number(_rand_range(add_r.get("km", {"min": 0.90, "max": 0.98}))),
        "h": _format_number(_rand_range(add_r.get("h", {"min": 3, "max": 6}))),
        "t_r": _format_number(_rand_range(add_r.get("t_r", {"min": 20, "max": 30}))),
        "e_n": _format_number(_rand_range(add_r.get("e_n", {"min": 0.08, "max": 0.12}))),
        "ts_e": _format_number(_rand_range(add_r.get("ts_e", {"min": 2.0, "max": 3.5}))),
    }

    conditions = {
        "oes": oes["name"],
        "theta_ohl": str(oes["theta_ohl"]),
    }

    # План: узлы A и B по бокам (по 2 клетки по x с каждой стороны), 1,2,3,4,5 — в центральной области.
    # A и B на разных сторонах. Масштаб подбирается так, чтобы расстояние A–B было 200–250 км.
    plan = config.get("table0_plan") or {"rows": 10, "cols": 12}
    n_rows = max(1, int(plan.get("rows", 10)))
    n_cols = max(1, int(plan.get("cols", 12)))
    min_dist = max(0, int(plan.get("min_circle_distance", 0)))
    grid_step_cm = max(0.1, float(plan.get("grid_step_cm", 1)))

    def _chebyshev_dist(r1, c1, r2, c2):
        return max(abs(r1 - r2), abs(c1 - c2))

    def _far_enough(r, c, placed):
        for (rp, cp) in placed:
            if _chebyshev_dist(r, c, rp, cp) < min_dist:
                return False
        return True

    def _min_dist_to_placed(r, c, placed):
        if not placed:
            return float("inf")
        return min(_chebyshev_dist(r, c, rp, cp) for (rp, cp) in placed)

    # Центральная область: по x от 2 до n_cols-3 (по 2 клетки слева и справа под A и B)
    center_col_min, center_col_max = 2, n_cols - 3
    if center_col_max < center_col_min:
        center_col_min, center_col_max = 0, n_cols - 1

    positions = []
    placed = []
    # Узел A — левая сторона (колонки 0 или 1), B — правая (n_cols-2 или n_cols-1), или наоборот
    side_a = random.choice(["left", "right"])
    side_b = "right" if side_a == "left" else "left"
    if side_a == "left":
        a_col = random.randint(0, min(1, n_cols - 1))
        b_col = random.choice([n_cols - 2, n_cols - 1]) if n_cols >= 2 else n_cols - 1
    else:
        a_col = random.choice([n_cols - 2, n_cols - 1]) if n_cols >= 2 else n_cols - 1
        b_col = random.randint(0, min(1, n_cols - 1))
    a_row = random.randint(0, n_rows - 1)
    b_row = random.randint(0, n_rows - 1)
    # Не совпадение A и B (разные стороны уже гарантированы при n_cols >= 4)
    placed.append((a_row, a_col))
    placed.append((b_row, b_col))
    positions.append({"row": a_row, "col": a_col, "label": "A"})
    positions.append({"row": b_row, "col": b_col, "label": "B"})

    center_cells = [
        (r, c) for r in range(n_rows) for c in range(center_col_min, center_col_max + 1)
        if (r, c) not in {(p[0], p[1]) for p in placed}
    ]
    random.shuffle(center_cells)
    # Разместить 1,2,3,4 в центральной области с min_circle_distance
    for label in ["1", "2", "3", "4"]:
        found = False
        for (r, c) in center_cells:
            if (r, c) in {(p[0], p[1]) for p in placed}:
                continue
            if _far_enough(r, c, placed):
                placed.append((r, c))
                positions.append({"row": r, "col": c, "label": label})
                found = True
                break
        if not found:
            for r in range(n_rows):
                for c in range(center_col_min, center_col_max + 1):
                    if (r, c) in {(p[0], p[1]) for p in placed}:
                        continue
                    if _far_enough(r, c, placed):
                        placed.append((r, c))
                        positions.append({"row": r, "col": c, "label": label})
                        found = True
                        break
                if found:
                    break
        if not found:
            best_r, best_c, best_d = None, None, -1
            for (r, c) in center_cells:
                if (r, c) in {(p[0], p[1]) for p in placed}:
                    continue
                d = _min_dist_to_placed(r, c, placed)
                if d > best_d:
                    best_d, best_r, best_c = d, r, c
            if best_r is not None:
                placed.append((best_r, best_c))
                positions.append({"row": best_r, "col": best_c, "label": label})

    # Узел 5 — в центре, с условием: расстояние от A до 5 ≈ от B до 5
    a_pos = (positions[0]["row"], positions[0]["col"])
    b_pos = (positions[1]["row"], positions[1]["col"])
    candidates_5 = [
        (r, c) for (r, c) in center_cells
        if (r, c) not in {(p[0], p[1]) for p in placed} and _far_enough(r, c, placed)
    ]
    if candidates_5:
        def _imbalance(r, c):
            d_a = _chebyshev_dist(r, c, a_pos[0], a_pos[1])
            d_b = _chebyshev_dist(r, c, b_pos[0], b_pos[1])
            return abs(d_a - d_b)
        candidates_5.sort(key=lambda rc: _imbalance(rc[0], rc[1]))
        r5, c5 = candidates_5[0]
        placed.append((r5, c5))
        positions.append({"row": r5, "col": c5, "label": "5"})
    else:
        # fallback: любая свободная центральная ячейка или середина по колонке
        mid_c = (center_col_min + center_col_max) // 2
        for r in range(n_rows):
            if (r, mid_c) not in {(p[0], p[1]) for p in placed} and _far_enough(r, mid_c, placed):
                placed.append((r, mid_c))
                positions.append({"row": r, "col": mid_c, "label": "5"})
                break
        else:
            for r in range(n_rows):
                for c in range(center_col_min, center_col_max + 1):
                    if (r, c) not in {(p[0], p[1]) for p in placed}:
                        placed.append((r, c))
                        positions.append({"row": r, "col": c, "label": "5"})
                        break
                else:
                    continue
                break

    # Масштаб: расстояние A–B не менее min_distance_ab_km и не более max_distance_ab_km (200–250 км).
    dist_ab_cells = _chebyshev_dist(a_pos[0], a_pos[1], b_pos[0], b_pos[1])
    dist_ab_cm = max(0.1, dist_ab_cells * grid_step_cm)
    min_distance_ab_km = max(1.0, float(plan.get("min_distance_ab_km", 200)))
    max_distance_ab_km = max(min_distance_ab_km, float(plan.get("max_distance_ab_km", 250)))
    scale_min = min_distance_ab_km / dist_ab_cm
    scale_max = max_distance_ab_km / dist_ab_cm
    if scale_max < scale_min:
        scale_max = scale_min + 1.0
    scale_km_per_cm = random.uniform(scale_min, scale_max)
    scale_text = str(round(scale_km_per_cm, 1) if scale_km_per_cm >= 10 else round(scale_km_per_cm, 2))

    return {
        "loads": loads,
        "generation": generation,
        "voltage_ps_a": voltage_ps_a,
        "voltage_table3": voltage_table3,
        "additional": additional,
        "conditions": conditions,
        "substation_positions": positions,
        "draw_substation_circles": config.get("draw_substation_circles", False),
        "circle_size_inches": config.get("circle_size_inches"),
        "scale_km_per_cm": round(scale_km_per_cm, 4),
        "scale_text": scale_text,
        "scale_cell": plan.get("scale_cell"),
    }


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))

    parser = argparse.ArgumentParser(
        description="Генерация N вариантов заданий со случайными значениями в заданных диапазонах"
    )
    parser.add_argument("--count", "-n", type=int, default=1,
                        help="Количество вариантов (по умолчанию 1)")
    parser.add_argument("--oes", "-o", default=None,
                        help="ОЭС для всех вариантов: Сибирь, Центр, Северо-запад, Юг, Восток, Урал. Если не указано — случайный выбор для каждого варианта.")
    parser.add_argument("--template", "-t", default=None,
                        help="Путь к шаблону .docx")
    parser.add_argument("--out-dir", "-d", default=None,
                        help="Папка для сохранения вариантов (по умолчанию — папка скрипта)")
    parser.add_argument("--config", "-c", default=None,
                        help="Путь к config_variants.json (ОЭС и диапазоны)")
    parser.add_argument("--seed", type=int, default=None,
                        help="Seed для ГПСЧ (для воспроизводимости)")
    parser.add_argument("--pdf", "-p", action="store_true",
                        help="Дополнительно сохранять каждый вариант в PDF")
    args = parser.parse_args()

    if args.count < 1:
        print("--count должен быть не меньше 1")
        return 1

    config_path = args.config or os.path.join(script_dir, "config_variants.json")
    if not os.path.isfile(config_path):
        print(f"Конфиг не найден: {config_path}")
        return 1
    config = load_config(config_path)

    template_path = args.template or os.path.join(script_dir, "Шаблон Район нагрузок электрической сети.docx")
    if not os.path.isfile(template_path):
        print(f"Шаблон не найден: {template_path}")
        return 1

    out_dir = args.out_dir or script_dir
    os.makedirs(out_dir, exist_ok=True)

    oes_names = [o["name"] for o in config.get("oes_list", [])]
    if args.oes:
        oes_normalized = (args.oes or "").strip().lower()
        match = next((n for n in oes_names if n.strip().lower() == oes_normalized), None)
        if not match:
            print(f"Неизвестная ОЭС: {args.oes}. Допустимые: {', '.join(oes_names)}")
            return 1
        args.oes = match

    for i in range(args.count):
        seed = (args.seed + i) if args.seed is not None else None
        data = generate_one_variant(config, oes_override=args.oes, seed=seed)
        data.setdefault("paragraph3", str(i + 1))  # номер в «на курсовой проект № …»
        doc = Document(template_path)
        fill_document(doc, data)
        out_name = f"variant_{i + 1:03d}.docx"
        out_path = os.path.join(out_dir, out_name)
        doc.save(out_path)
        line = f"  {out_name}  OES: {data['conditions']['oes']}, theta_ohl: {data['conditions']['theta_ohl']} C"
        if args.pdf:
            pdf_path = convert_docx_to_pdf(out_path)
            if pdf_path:
                line += f"  → {os.path.basename(pdf_path)}"
        print(line)

    if args.pdf:
        print("Для PDF нужны: pip install docx2pdf и Microsoft Word или LibreOffice (soffice).")
    print(f"Создано вариантов: {args.count} в папке {out_dir}")
    return 0


if __name__ == "__main__":
    exit(main() or 0)
