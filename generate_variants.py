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

from fill_template import fill_document


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

    gen_r = ranges.get("generation", {})
    gen_node_idx = random.randint(0, n_loads - 1) if n_loads else 0
    gen_node = loads[gen_node_idx]["number_ps"] if loads else "ПС-1"
    generation = {
        "gen_node": gen_node,
        "p_ust": _format_number(_rand_range(gen_r.get("p_ust", {"min": 5, "max": 25}))),
        "gen_count": _format_number(_rand_range(gen_r.get("gen_count", {"min": 1, "max": 4}))),
    }

    v_r = ranges.get("voltage_ps_a", {})
    voltage_ps_a = {
        "u_max": _format_number(_rand_range(v_r.get("u_max", {"min": 114, "max": 118}))),
        "u_min": _format_number(_rand_range(v_r.get("u_min", {"min": 108, "max": 114}))),
    }

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

    # Координаты на плане (таблица 0): 5 кружков — ПС А (первый), затем подстанции 1–4.
    # Условие: расстояние между любыми двумя кружками не менее min_circle_distance ячеек (метрика Чебышёва).
    plan = config.get("table0_plan") or {"rows": 10, "cols": 10}
    n_rows = max(1, int(plan.get("rows", 10)))
    n_cols = max(1, int(plan.get("cols", 10)))
    min_dist = max(0, int(plan.get("min_circle_distance", 0)))

    def _chebyshev_dist(r1, c1, r2, c2):
        return max(abs(r1 - r2), abs(c1 - c2))

    def _far_enough(r, c, placed):
        """Новая точка (r, c) не ближе min_dist ни к одной из уже размещённых (по Чебышёву)."""
        for (rp, cp) in placed:
            if _chebyshev_dist(r, c, rp, cp) < min_dist:
                return False
        return True

    def _min_dist_to_placed(r, c, placed):
        """Минимальное расстояние от (r, c) до размещённых точек (Чебышёв)."""
        if not placed:
            return float("inf")
        return min(_chebyshev_dist(r, c, rp, cp) for (rp, cp) in placed)

    labels = ["ПС А"] + [loads[i]["number_ps"] for i in range(min(4, len(loads)))]
    n_circles = 5
    positions = []
    placed = []
    # Расположение ПС А (первый кружок): только слева или справа на сетке
    ps_a_side = (plan.get("ps_a_side") or "random").strip().lower()
    if ps_a_side == "random":
        ps_a_side = random.choice(["left", "right"])
    if ps_a_side == "left":
        ps_a_col = 0
    else:
        ps_a_col = n_cols - 1
    ps_a_row = random.randint(0, n_rows - 1)
    placed.append((ps_a_row, ps_a_col))
    positions.append({"row": ps_a_row, "col": ps_a_col, "label": labels[0]})

    for idx in range(1, n_circles):
        label = labels[idx] if idx < len(labels) else str(idx + 1)
        for _ in range(200):
            r = random.randint(0, n_rows - 1)
            c = random.randint(0, n_cols - 1)
            if (r, c) not in {(p[0], p[1]) for p in placed} and _far_enough(r, c, placed):
                placed.append((r, c))
                positions.append({"row": r, "col": c, "label": label})
                break
        else:
            # Fallback: ищем любую свободную ячейку с допустимым расстоянием; если нет — ставим в точку, наиболее удалённую от уже размещённых
            found = False
            for r in range(n_rows):
                for c in range(n_cols):
                    if (r, c) not in {(p[0], p[1]) for p in placed} and _far_enough(r, c, placed):
                        placed.append((r, c))
                        positions.append({"row": r, "col": c, "label": label})
                        found = True
                        break
                if found:
                    break
            if not found:
                best_r, best_c = None, None
                best_d = -1
                for r in range(n_rows):
                    for c in range(n_cols):
                        if (r, c) in {(p[0], p[1]) for p in placed}:
                            continue
                        d = _min_dist_to_placed(r, c, placed)
                        if d > best_d:
                            best_d, best_r, best_c = d, r, c
                if best_r is not None:
                    placed.append((best_r, best_c))
                    positions.append({"row": best_r, "col": best_c, "label": label})
                else:
                    r_fb = len(positions) % n_rows
                    c_fb = (len(positions) * 2) % n_cols
                    placed.append((r_fb, c_fb))
                    positions.append({"row": r_fb, "col": c_fb, "label": label})

    # Масштаб плана: шаг сетки grid_step_cm см; макс. расстояние от ПС А до дальней подстанции — не более max_distance_km км.
    # scale_km_per_cm: 1 см на плане = X км в натуре (подбирается так, чтобы макс. расстояние по плану давало ровно max_distance_km).
    grid_step_cm = max(0.1, float(plan.get("grid_step_cm", 1)))
    max_distance_km = max(1.0, float(plan.get("max_distance_km", 200)))
    r0, c0 = positions[0]["row"], positions[0]["col"]
    max_dist_cells = 0
    for pos in positions[1:]:
        d = _chebyshev_dist(r0, c0, pos["row"], pos["col"])
        if d > max_dist_cells:
            max_dist_cells = d
    max_dist_cm = max_dist_cells * grid_step_cm
    if max_dist_cm < 0.1:
        max_dist_cm = 0.1
    scale_km_per_cm = max_distance_km / max_dist_cm
    scale_text = str(round(scale_km_per_cm, 1) if scale_km_per_cm >= 10 else round(scale_km_per_cm, 2))

    return {
        "loads": loads,
        "generation": generation,
        "voltage_ps_a": voltage_ps_a,
        "additional": additional,
        "conditions": conditions,
        "substation_positions": positions,
        "draw_substation_circles": config.get("draw_substation_circles", False),
        "circle_size_inches": config.get("circle_size_inches"),
        "scale_km_per_cm": round(scale_km_per_cm, 4),
        "scale_text": scale_text,
        "max_distance_cells": max_dist_cells,
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
        print(f"  {out_name}  OES: {data['conditions']['oes']}, theta_ohl: {data['conditions']['theta_ohl']} C")

    print(f"Создано вариантов: {args.count} в папке {out_dir}")
    return 0


if __name__ == "__main__":
    exit(main() or 0)
