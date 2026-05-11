import sys
import os
import csv
import re
import math
from collections import Counter, defaultdict
from datetime import datetime
from statistics import median

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# PIN POSITION SIMPLE VIEW V6
# Widok prosty + możliwość ustawienia punktu nominalnego i tolerancji
# osobno dla każdego pina i każdej osi.
#
# Raport:
# Pozycja pinów | Odchylenie pinów | Oś X | Oś Y | Oś Z | Średnia odchyłka od punktu [mm]
# ============================================================

IGNORE_ZERO_VALUES = True
MIN_VALID_VALUES_FOR_AUTO_TARGET = 10

# Ile najbardziej problematycznych pinów pokazać w kolumnie „Odchylenie pinów”.
TOP_BAD_PINS_PER_FILE = 4

# ============================================================
# TOLERANCJE WG WYROBU
# XY oznacza: X i Y mają tę samą tolerancję.
# Wszystkie wartości są w mm.
# ============================================================

PRODUCT_TOLERANCE_BY_REF = {
    "GSR":   {"X": 0.30, "Y": 0.30, "Z": 0.25},
    "RN2":   {"X": 0.30, "Y": 0.30, "Z": 0.30},
    "RN1":   {"X": 0.30, "Y": 0.30, "Z": 0.25},
    "FCA":   {"X": 0.30, "Y": 0.30, "Z": 0.25},
    "SWEET": {"X": 0.25, "Y": 0.25, "Z": 0.25},
    "HKMC":  {"X": 0.30, "Y": 0.30, "Z": 0.25},
}

# Jeśli produkt nie zostanie rozpoznany, program użyje tych wartości awaryjnych.
DEFAULT_TOLERANCE_BY_AXIS = {
    "X": 0.30,
    "Y": 0.30,
    "Z": 0.25,
}

# Jeśli dla pina/osi nie wpiszesz punktu nominalnego,
# program policzy punkt automatycznie jako medianę z danych.
# Czyli najlepiej:
# - jeśli znasz punkt nominalny z procesu/rysunku -> wpisz go w PIN_SETTINGS_BY_REF,
# - jeśli nie znasz -> zostaw pusty, program użyje mediany.
USE_MEDIAN_AS_TARGET_IF_NOT_SET = True

# Zakres zabezpieczający przed ewidentnymi śmieciami w danych.
# To NIE jest tolerancja procesu.
GROSS_MIN = -1.0
GROSS_MAX = 25.0

# ============================================================
# MAPA RZĘDÓW PINÓW
# Tutaj ustawiasz fizyczny układ pinów.
# Jeśli produkt nie jest wpisany poniżej, program zrobi układ automatyczny.
# ============================================================

PIN_LAYOUTS_BY_REF = {
    "GSR": {
        "ROW_1": [1, 5],
        "ROW_2": [2, 6],
        "ROW_3": [3, 7],
        "ROW_4": [4, 8],
    },

    "SWEET": {
        "ROW_1": [1, 5],
        "ROW_2": [2, 6],
        "ROW_3": [3, 7],
        "ROW_4": [4, 8],
        "ROW_5": [9],
    },

    "HKMC": {
        "ROW_1": [1, 9],
        "ROW_2": [2, 10],
        "ROW_3": [3, 11],
        "ROW_4": [4, 12],
        "ROW_5": [5, 13],
        "ROW_6": [6, 14],
        "ROW_7": [7, 15],
        "ROW_8": [8, 16],
    },
}

AUTO_LAYOUT_MODE = "TWO_COLUMNS"

# ============================================================
# PUNKTY NOMINALNE DLA PINÓW
# ============================================================

def make_gsr_pin_settings():
    """
    GSR:
    Pin 1: X = 5.085, Y = 11.063, Z = 6.209
    Pin 2-4: Y = Y1 - n * 1.814
    Pin 5-8: X = 3.700, te same wartości Y co Pin1-4

    Założenie:
    Z = 6.209 dla wszystkich pinów GSR.
    Jeśli Z dla pinów 2-8 jest inne, trzeba wpisać osobno.
    """

    x_left = 5.085
    x_right = 3.700
    y1 = 11.063
    y_step = 1.814
    z_all = 6.209

    settings = {}

    # Piny 1-4
    for pin in range(1, 5):
        n = pin - 1
        y = y1 - n * y_step

        settings[pin] = {
            "target": {
                "X": x_left,
                "Y": round(y, 3),
                "Z": z_all,
            }
        }

    # Piny 5-8
    for pin in range(5, 9):
        base_pin = pin - 4
        n = base_pin - 1
        y = y1 - n * y_step

        settings[pin] = {
            "target": {
                "X": x_right,
                "Y": round(y, 3),
                "Z": z_all,
            }
        }

    return settings


# ============================================================
# PUNKTY NOMINALNE I TOLERANCJE DLA KAŻDEGO PINA
# ============================================================
#
# To jest najważniejsze miejsce do edycji.
# Możesz ustawić osobny punkt nominalny oraz osobną tolerancję dla:
# - produktu, np. GSR/SWEET/HKMC,
# - pina, np. 1, 2, 3,
# - osi, X/Y/Z.
#
# Jeśli nie wpiszesz target dla danego pina/osi, program użyje mediany z pliku.
# Jeśli nie wpiszesz tolerance dla danego pina/osi, program użyje tolerancji produktu.
#
# Przykład pełny dla Pin1:
# "GSR": {
#     1: {
#         "target": {"X": 2.850, "Y": 3.450, "Z": 5.650},
#         "tol":    {"X": 0.100, "Y": 0.100, "Z": 0.050},
#     },
# }
#
# Przykład: tylko zmiana tolerancji dla Pin2 na osi Z:
# "GSR": {
#     2: {
#         "tol": {"Z": 0.030},
#     },
# }
# ============================================================

PIN_SETTINGS_BY_REF = {
    "GSR": make_gsr_pin_settings(),

    "RN2": {
        # Tutaj dopisz punkty nominalne RN2, gdy je podasz.
        # Przykład:
        # 1: {
        #     "target": {"X": 0.000, "Y": 0.000, "Z": 0.000},
        # },
    },

    "RN1": {
        # Tutaj dopisz punkty nominalne RN1.
    },

    "FCA": {
        # Tutaj dopisz punkty nominalne FCA.
    },

    "SWEET": {
        # Tutaj dopisz punkty nominalne SWEET.
    },

    "HKMC": {
        # Tutaj dopisz punkty nominalne HKMC.
    },
}

# ============================================================
# FUNKCJE POMOCNICZE
# ============================================================

def get_exe_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def clean_cell(value):
    if value is None:
        return ""
    return str(value).strip().strip('"')


def parse_float(value):
    text = clean_cell(value).replace(",", ".")
    if text == "" or text.lower() in ["nan", "none", "null", "-"]:
        return None
    try:
        num = float(text)
        if math.isnan(num) or math.isinf(num):
            return None
        return num
    except Exception:
        return None


def detect_dialect(file_path):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore", newline="") as f:
            sample = f.read(12000)
            return csv.Sniffer().sniff(sample)
    except Exception:
        class SemiDialect(csv.excel):
            delimiter = ";"
        return SemiDialect


def pin_axis_from_metric(metric):
    m = re.search(r"Data[_ ]Pin(\d+)[_ ]Position([XYZ])", metric, re.IGNORECASE)
    if not m:
        return None, None
    return int(m.group(1)), m.group(2).upper()


def is_position_metric(metric):
    pin, axis = pin_axis_from_metric(metric)
    return pin is not None and axis is not None


def detect_ref_from_file(file_base):
    upper = file_base.upper()
    for key in set(list(PIN_LAYOUTS_BY_REF.keys()) + list(PIN_SETTINGS_BY_REF.keys())):
        if key in upper:
            return key
    return os.path.splitext(file_base)[0].split("_")[0].upper()


def median_or_none(values):
    vals = sorted([v for v in values if v is not None])
    if not vals:
        return None
    return median(vals)


def pin_boxes(pins, bad_pins):
    boxes = []
    for p in pins:
        if p in bad_pins:
            boxes.append(f"[{p}]")
        else:
            boxes.append("[ ]")
    return " ".join(boxes)


def position_boxes(pins):
    return " ".join(f"[{p}]" for p in pins)


def fmt_mm(value):
    if value is None:
        return "-"
    return f"{value:.3f}"


def get_pin_setting(ref, pin, axis, setting_name):
    """setting_name = 'target' albo 'tol'."""
    ref_upper = (ref or "").upper()

    # Najpierw próbujemy dokładnie po ref.
    if ref_upper in PIN_SETTINGS_BY_REF:
        pin_cfg = PIN_SETTINGS_BY_REF.get(ref_upper, {}).get(pin, {})
        axis_cfg = pin_cfg.get(setting_name, {})
        if axis in axis_cfg:
            return axis_cfg[axis]

    # Potem próbujemy po kluczu zawartym w ref, np. ref = GSR_11_05, key = GSR.
    for key, product_cfg in PIN_SETTINGS_BY_REF.items():
        if key.upper() in ref_upper:
            pin_cfg = product_cfg.get(pin, {})
            axis_cfg = pin_cfg.get(setting_name, {})
            if axis in axis_cfg:
                return axis_cfg[axis]

    return None


def get_product_default_tolerance(ref, axis):
    """
    Pobiera tolerancję domyślną dla wyrobu.
    Przykład:
    GSR -> X/Y = 0.30, Z = 0.25
    SWEET -> X/Y/Z = 0.25
    """
    ref_upper = (ref or "").upper()

    for product_name, tol_cfg in PRODUCT_TOLERANCE_BY_REF.items():
        if product_name.upper() in ref_upper:
            return tol_cfg.get(axis, DEFAULT_TOLERANCE_BY_AXIS.get(axis))

    return DEFAULT_TOLERANCE_BY_AXIS.get(axis)


def avg_text_for_pins(pins, avg_abs_delta_by_pin_axis, target_by_pin_axis, tol_by_pin_axis):
    """
    Przykład:
    Pin[1]: x=0.012/tol=0.300/pkt=5.085, y=0.003/tol=0.300/pkt=11.063, z=0.044/tol=0.250/pkt=6.209 mm
    """
    parts = []
    for p in pins:
        axis_parts = []
        for axis in ["X", "Y", "Z"]:
            avg_dev = avg_abs_delta_by_pin_axis.get((p, axis), 0.0)
            tol = tol_by_pin_axis.get((p, axis))
            target = target_by_pin_axis.get((p, axis))
            axis_parts.append(
                f"{axis.lower()}={fmt_mm(avg_dev)} / tol={fmt_mm(tol)} / pkt={fmt_mm(target)}"
            )
        parts.append(f"Pin[{p}]: " + ", ".join(axis_parts) + " mm")
    return " | ".join(parts)

# ============================================================
# CZYTANIE CSV
# ============================================================

def read_rows(file_path):
    dialect = detect_dialect(file_path)
    with open(file_path, "r", encoding="utf-8", errors="ignore", newline="") as f:
        return [row for row in csv.reader(f, dialect=dialect)]


def find_header(rows):
    for i, row in enumerate(rows):
        normalized = [clean_cell(c).upper() for c in row]
        if "NAME" in normalized:
            name_col = normalized.index("NAME")
            data_start = None
            for j, cell in enumerate(normalized):
                if cell == "USL (CALCULATED)":
                    data_start = j + 1
                    break
            if data_start is None:
                data_start = name_col + 16
            return i, name_col, data_start
    raise ValueError("Nie znaleziono wiersza z kolumną Name")


def collect_metrics(rows, header_idx, name_col):
    metrics = {}
    raw_rows = []
    for row in rows[header_idx + 1:]:
        name = clean_cell(row[name_col]) if len(row) > name_col else ""
        raw_rows.append((name, row))
        if name:
            metrics[name] = row
    return metrics, raw_rows


def extract_records(file_path):
    file_base = os.path.basename(file_path)
    ref = detect_ref_from_file(file_base)

    rows = read_rows(file_path)
    header_idx, name_col, data_start = find_header(rows)
    metrics, raw_rows = collect_metrics(rows, header_idx, name_col)

    position_metrics = [m for m in metrics if is_position_metric(m)]
    if not position_metrics:
        raise ValueError("Nie znaleziono metryk Data_PinN_PositionX/Y/Z")

    max_len = max(len(row) for _, row in raw_rows) if raw_rows else 0

    records = []
    for col in range(data_start, max_len):
        record = {
            "file": file_base,
            "ref": ref,
            "sample_index": col - data_start + 1,
            "positions": {},
        }

        for metric in position_metrics:
            row = metrics[metric]
            pin, axis = pin_axis_from_metric(metric)
            raw = row[col] if col < len(row) else ""
            val = parse_float(raw)
            record["positions"][(pin, axis)] = val

        records.append(record)

    return records

# ============================================================
# LAYOUT PINÓW
# ============================================================

def make_auto_two_column_layout(active_pins):
    pins = sorted(active_pins)
    rows = {}
    if not pins:
        return rows

    half = (len(pins) + 1) // 2
    left = pins[:half]
    right = pins[half:]

    for i, left_pin in enumerate(left):
        row_pins = [left_pin]
        if i < len(right):
            row_pins.append(right[i])
        rows[f"ROW_{i + 1}"] = row_pins

    return rows


def get_layout_for_ref(ref, active_pins):
    ref_upper = (ref or "").upper()
    active = set(active_pins)

    for key, layout in PIN_LAYOUTS_BY_REF.items():
        if key.upper() in ref_upper:
            cleaned_layout = {
                row: [p for p in pins if p in active]
                for row, pins in layout.items()
                if any(p in active for p in pins)
            }
            if cleaned_layout:
                return cleaned_layout

    if AUTO_LAYOUT_MODE == "TWO_COLUMNS":
        return make_auto_two_column_layout(active)

    return {f"ROW_{i + 1}": [pin] for i, pin in enumerate(sorted(active))}

# ============================================================
# ANALIZA
# ============================================================

def build_auto_targets(records):
    values = defaultdict(list)
    active_pins = defaultdict(set)

    for r in records:
        group = (r["file"], r["ref"])
        for (pin, axis), val in r["positions"].items():
            if val is None:
                continue
            if IGNORE_ZERO_VALUES and val == 0:
                continue
            if val < GROSS_MIN or val > GROSS_MAX:
                continue
            values[(r["file"], r["ref"], pin, axis)].append(val)
            active_pins[group].add(pin)

    auto_target = {}
    for key, vals in values.items():
        if len(vals) >= MIN_VALID_VALUES_FOR_AUTO_TARGET:
            auto_target[key] = median_or_none(vals)

    return auto_target, active_pins


def analyze(records):
    auto_target, active_pins = build_auto_targets(records)
    result = {}

    for group, pins in active_pins.items():
        file, ref = group
        layout = get_layout_for_ref(ref, pins)

        bad_by_axis = {"X": set(), "Y": set(), "Z": set()}
        max_abs_delta_by_pin = defaultdict(float)
        sum_abs_delta_by_pin_axis = defaultdict(float)
        count_delta_by_pin_axis = Counter()
        target_by_pin_axis = {}
        tol_by_pin_axis = {}

        # Przygotowanie target/tolerance dla każdego pina i osi.
        for pin in pins:
            for axis in ["X", "Y", "Z"]:
                configured_target = get_pin_setting(ref, pin, axis, "target")
                configured_tol = get_pin_setting(ref, pin, axis, "tol")

                if configured_target is not None:
                    target = configured_target
                elif USE_MEDIAN_AS_TARGET_IF_NOT_SET:
                    target = auto_target.get((file, ref, pin, axis))
                else:
                    target = None

                if configured_tol is not None:
                    tol = configured_tol
                else:
                    tol = get_product_default_tolerance(ref, axis)

                target_by_pin_axis[(pin, axis)] = target
                tol_by_pin_axis[(pin, axis)] = tol

        # Analiza pomiarów względem punktu nominalnego i tolerancji.
        for r in records:
            if (r["file"], r["ref"]) != group:
                continue

            for (pin, axis), value in r["positions"].items():
                if value is None:
                    continue
                if IGNORE_ZERO_VALUES and value == 0:
                    continue
                if value < GROSS_MIN or value > GROSS_MAX:
                    continue

                target = target_by_pin_axis.get((pin, axis))
                tol = tol_by_pin_axis.get((pin, axis))

                if target is None or tol is None:
                    continue

                abs_delta = abs(value - target)
                sum_abs_delta_by_pin_axis[(pin, axis)] += abs_delta
                count_delta_by_pin_axis[(pin, axis)] += 1

                if abs_delta > tol:
                    bad_by_axis[axis].add(pin)
                    max_abs_delta_by_pin[pin] = max(max_abs_delta_by_pin[pin], abs_delta)

        avg_abs_delta_by_pin_axis = {}
        for key, total in sum_abs_delta_by_pin_axis.items():
            count = count_delta_by_pin_axis[key]
            avg_abs_delta_by_pin_axis[key] = total / count if count else 0.0

        top_bad_pins = set(
            pin for pin, _delta in sorted(
                max_abs_delta_by_pin.items(),
                key=lambda x: x[1],
                reverse=True
            )[:TOP_BAD_PINS_PER_FILE]
        )

        result[group] = {
            "layout": layout,
            "bad_overall": top_bad_pins,
            "bad_by_axis": bad_by_axis,
            "avg_abs_delta_by_pin_axis": avg_abs_delta_by_pin_axis,
            "target_by_pin_axis": target_by_pin_axis,
            "tol_by_pin_axis": tol_by_pin_axis,
        }

    return result

# ============================================================
# EXCEL — PROSTY WIDOK
# ============================================================

def auto_width(ws):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = 10
        for cell in col:
            if cell.value is not None:
                width = max(width, min(95, len(str(cell.value)) + 4))
        ws.column_dimensions[letter].width = width


def style_simple_sheet(ws):
    blue = PatternFill("solid", fgColor="1F4E78")
    white_font = Font(color="FFFFFF", bold=True)
    red = PatternFill("solid", fgColor="F8696B")
    yellow = PatternFill("solid", fgColor="FFEB84")
    green = PatternFill("solid", fgColor="E2F0D9")
    grey = PatternFill("solid", fgColor="D9EAD3")
    thin = Side(style="thin", color="BFBFBF")

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in range(1, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, str) and value.startswith("PLIK:"):
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = green
                ws.cell(row=row, column=col).font = Font(bold=True)

    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == "Pozycja pinów":
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = blue
                ws.cell(row=row, column=col).font = white_font

    for row in range(1, ws.max_row + 1):
        for col in range(2, 6):
            value = ws.cell(row=row, column=col).value
            if isinstance(value, str) and re.search(r"\[\d+\]", value):
                ws.cell(row=row, column=col).fill = red
                ws.cell(row=row, column=col).font = Font(bold=True)
            elif isinstance(value, str) and "[ ]" in value:
                ws.cell(row=row, column=col).fill = yellow

        value = ws.cell(row=row, column=6).value
        if isinstance(value, str) and "Pin[" in value:
            ws.cell(row=row, column=6).fill = grey

    auto_width(ws)
    ws.freeze_panes = "A1"


def create_excel(output_file, analysis_result):
    wb = Workbook()
    ws = wb.active
    ws.title = "Widok_Pinow"

    current_row = 1

    for (file, ref), data in sorted(analysis_result.items()):
        layout = data["layout"]
        bad_overall = data["bad_overall"]
        bad_by_axis = data["bad_by_axis"]
        avg_abs_delta_by_pin_axis = data["avg_abs_delta_by_pin_axis"]
        target_by_pin_axis = data["target_by_pin_axis"]
        tol_by_pin_axis = data["tol_by_pin_axis"]

        ws.cell(row=current_row, column=1, value=f"PLIK: {file}")
        ws.cell(row=current_row, column=2, value=f"REF: {ref}")
        ws.cell(row=current_row, column=3, value=f"TOP pinów: {TOP_BAD_PINS_PER_FILE}")
        current_row += 1

        headers = [
            "Pozycja pinów",
            "Odchylenie pinów",
            "Oś X",
            "Oś Y",
            "Oś Z",
            "Średnia odchyłka / tolerancja / punkt [mm]",
        ]
        for col, header in enumerate(headers, start=1):
            ws.cell(row=current_row, column=col, value=header)
        current_row += 1

        for row_name, pins in layout.items():
            ws.cell(row=current_row, column=1, value=position_boxes(pins))
            ws.cell(row=current_row, column=2, value=pin_boxes(pins, bad_overall))
            ws.cell(row=current_row, column=3, value=pin_boxes(pins, bad_by_axis["X"]))
            ws.cell(row=current_row, column=4, value=pin_boxes(pins, bad_by_axis["Y"]))
            ws.cell(row=current_row, column=5, value=pin_boxes(pins, bad_by_axis["Z"]))
            ws.cell(
                row=current_row,
                column=6,
                value=avg_text_for_pins(
                    pins,
                    avg_abs_delta_by_pin_axis,
                    target_by_pin_axis,
                    tol_by_pin_axis,
                )
            )
            current_row += 1

        current_row += 2

    style_simple_sheet(ws)
    wb.save(output_file)

# ============================================================
# KONSOLA
# ============================================================

def print_console_view(analysis_result):
    print("\n" + "=" * 130)
    print("PROSTY WIDOK ODCHYLEŃ PINÓW - TARGET/TOLERANCE")
    print("=" * 130)

    for (file, ref), data in sorted(analysis_result.items()):
        print(f"\nPLIK: {file} | REF: {ref}")
        print("-" * 130)
        print(
            f"{'Pozycja pinów':<18} | "
            f"{'Odchylenie pinów':<20} | "
            f"{'Oś X':<12} | "
            f"{'Oś Y':<12} | "
            f"{'Oś Z':<12} | "
            f"Średnia odchyłka / tolerancja / punkt [mm]"
        )
        print("-" * 130)

        layout = data["layout"]
        bad_overall = data["bad_overall"]
        bad_by_axis = data["bad_by_axis"]
        avg_abs_delta_by_pin_axis = data["avg_abs_delta_by_pin_axis"]
        target_by_pin_axis = data["target_by_pin_axis"]
        tol_by_pin_axis = data["tol_by_pin_axis"]

        for row_name, pins in layout.items():
            pos = position_boxes(pins)
            overall = pin_boxes(pins, bad_overall)
            axis_x = pin_boxes(pins, bad_by_axis["X"])
            axis_y = pin_boxes(pins, bad_by_axis["Y"])
            axis_z = pin_boxes(pins, bad_by_axis["Z"])
            avg_txt = avg_text_for_pins(
                pins,
                avg_abs_delta_by_pin_axis,
                target_by_pin_axis,
                tol_by_pin_axis,
            )
            print(f"{pos:<18} | {overall:<20} | {axis_x:<12} | {axis_y:<12} | {axis_z:<12} | {avg_txt}")

        print("-" * 130)

# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 2:
        print("Przeciągnij jeden lub więcej plików CSV na program EXE.")
        input("ENTER aby wyjść...")
        return

    csv_files = [
        p for p in sys.argv[1:]
        if p.lower().endswith(".csv") and os.path.isfile(p)
    ]

    if not csv_files:
        print("Nie znaleziono plików CSV.")
        input("ENTER aby wyjść...")
        return

    all_records = []

    print("Analizowane pliki:")
    for file_path in csv_files:
        print("-", file_path)
        try:
            records = extract_records(file_path)
            all_records.extend(records)
            print(f"  OK: {len(records)} rekordów")
        except Exception as e:
            print(f"  BŁĄD: {e}")

    if not all_records:
        print("Brak danych do analizy.")
        input("ENTER aby wyjść...")
        return

    analysis_result = analyze(all_records)
    print_console_view(analysis_result)

    out_dir = get_exe_folder()
    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    output_file = os.path.join(out_dir, f"pin_target_tolerance_view__{timestamp}.xlsx")

    create_excel(output_file, analysis_result)

    print("\nZapisano prosty raport Excel:")
    print(output_file)

    input("\nENTER aby zamknąć...")


if __name__ == "__main__":
    main()
