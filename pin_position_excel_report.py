
import sys
import os
import csv
import re
import math
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from statistics import median

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

# ============================================================
# KONFIGURACJA UŻYTKOWNIKA
# ============================================================

ROBUST_Z_THRESHOLD = 5.0
MIN_VALID_VALUES_FOR_ANALYSIS = 10
IGNORE_ZERO_VALUES = True
ANALYZE_FAIL_ONLY = False       # True = analizuje tylko StatusBits != 0 albo Failure_Bits != 0
MAX_DEVIATIONS_IN_EXCEL = 200000
TOP_N = 15

# Jeżeli MAD=0 lub bardzo mały, używany jest próg absolutny.
FALLBACK_TOLERANCE = {
    "X": 0.15,
    "Y": 0.15,
    "Z": 0.08,
}

# Granice bardzo grubych błędów pomiaru, np. -0.03, 28.474 itp.
GROSS_MIN = -1.0
GROSS_MAX = 25.0

# Mapa rzędów pinów. AUTO = program sam stworzy pary zależnie od liczby aktywnych pinów.
PIN_ROWS_MODE = "AUTO"
CUSTOM_PIN_ROWS = {
    # Przykład dla GSR, jeśli fizycznie tak wygląda:
    # "ROW_1": [1, 5],
    # "ROW_2": [2, 6],
    # "ROW_3": [3, 7],
    # "ROW_4": [4, 8],
}
MIN_PINS_IN_ROW_FAIL = 2

# ============================================================
# FUNKCJE POMOCNICZE
# ============================================================

def get_exe_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def detect_dialect(file_path):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore", newline="") as f:
            sample = f.read(12000)
            return csv.Sniffer().sniff(sample)
    except Exception:
        class SemiDialect(csv.excel):
            delimiter = ";"
        return SemiDialect


def clean_cell(value):
    if value is None:
        return ""
    return str(value).strip().strip('"')


def parse_float(value):
    text = clean_cell(value).replace(",", ".")
    if text == "" or text.lower() in ["nan", "none", "null", "-"]:
        return None
    try:
        number = float(text)
        if math.isnan(number) or math.isinf(number):
            return None
        return number
    except Exception:
        return None


def parse_datetime_from_text(text):
    text = clean_cell(text)
    if not text:
        return None

    formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%m/%d/%Y, %I:%M:%S %p",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y, %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            pass

    patterns = [
        r"\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}(:\d{2})?",
        r"\d{2}\.\d{2}\.\d{4}\s+\d{1,2}:\d{2}(:\d{2})?",
        r"\d{1,2}/\d{1,2}/\d{4},?\s+\d{1,2}:\d{2}:\d{2}\s*(AM|PM)?",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            found = m.group(0)
            for fmt in formats:
                try:
                    return datetime.strptime(found, fmt)
                except Exception:
                    pass
    return None


def hour_bucket(dt):
    if dt is None:
        return "UNKNOWN_TIME"
    return dt.strftime("%Y-%m-%d %H:00")


def pin_axis_from_name(metric_name):
    m = re.search(r"Data[_ ]Pin(\d+)[_ ]Position([XYZ])", metric_name, re.IGNORECASE)
    if not m:
        return None, None
    return int(m.group(1)), m.group(2).upper()


def is_position_metric(metric_name):
    p, a = pin_axis_from_name(metric_name)
    return p is not None and a is not None


def robust_stats(values):
    vals = sorted([v for v in values if v is not None])
    if not vals:
        return None, None
    med = median(vals)
    mad = median([abs(v - med) for v in vals])
    return med, mad


def robust_z(value, med, mad):
    if value is None or med is None or mad is None or mad == 0:
        return None
    return 0.6745 * (value - med) / mad


def normalize_ref_name(ref):
    ref = clean_cell(ref)
    return ref if ref else "UNKNOWN_REF"

# ============================================================
# PARSER CSV TRANSPOSED
# ============================================================

def read_rows(file_path):
    dialect = detect_dialect(file_path)
    with open(file_path, "r", encoding="utf-8", errors="ignore", newline="") as f:
        return [row for row in csv.reader(f, dialect=dialect)]


def find_header_info(rows):
    """Zwraca: header_idx, name_col, data_start_col."""
    for i, row in enumerate(rows):
        norm = [clean_cell(c).upper() for c in row]
        if "NAME" in norm:
            name_col = norm.index("NAME")
            data_start_col = None
            for j, c in enumerate(norm):
                if c == "USL (CALCULATED)":
                    data_start_col = j + 1
                    break
            if data_start_col is None:
                data_start_col = name_col + 16
            return i, name_col, data_start_col
    raise ValueError("Nie znaleziono wiersza nagłówka z kolumną Name")


def collect_named_rows(rows, header_idx, name_col):
    metrics = {}
    raw_after_header = []
    for row in rows[header_idx + 1:]:
        name = clean_cell(row[name_col]) if len(row) > name_col else ""
        raw_after_header.append((name, row))
        if name:
            metrics[name] = row
    return metrics, raw_after_header


def find_datetime_values(raw_after_header, data_start_col):
    best_row = None
    best_count = 0
    for name, row in raw_after_header:
        count = 0
        for val in row[data_start_col:]:
            if parse_datetime_from_text(val):
                count += 1
        if count > best_count:
            best_count = count
            best_row = row
    if best_count == 0:
        return None
    return best_row


def extract_records(file_path):
    rows = read_rows(file_path)
    header_idx, name_col, data_start_col = find_header_info(rows)
    header = rows[header_idx]
    metrics, raw_after_header = collect_named_rows(rows, header_idx, name_col)

    position_metrics = [m for m in metrics if is_position_metric(m)]
    if not position_metrics:
        raise ValueError("Brak metryk Data_PinN_PositionX/Y/Z")

    max_len = max(len(r) for _, r in raw_after_header) if raw_after_header else len(header)
    dt_row = find_datetime_values(raw_after_header, data_start_col)

    status_row = None
    for key in ["StatusBits", "Status_Bits"]:
        if key in metrics:
            status_row = metrics[key]
            break

    failure_row = None
    for key in ["Failure_Bits", "FailureBits"]:
        if key in metrics:
            failure_row = metrics[key]
            break

    ref_row = metrics.get("Ref_Name") or metrics.get("RefName")

    file_base = os.path.basename(file_path)
    records = []

    for col in range(data_start_col, max_len):
        sample_id = clean_cell(header[col]) if col < len(header) else ""
        dt = parse_datetime_from_text(dt_row[col]) if dt_row and col < len(dt_row) else None
        status = parse_float(status_row[col]) if status_row and col < len(status_row) else None
        failure = parse_float(failure_row[col]) if failure_row and col < len(failure_row) else None
        ref_name = normalize_ref_name(ref_row[col] if ref_row and col < len(ref_row) else "")

        # jeśli ref_name pusty w pomiarach, bierz z nazwy pliku
        if ref_name == "UNKNOWN_REF":
            ref_name = os.path.splitext(file_base)[0].split("_")[0]

        record = {
            "file": file_base,
            "file_path": file_path,
            "ref_name": ref_name,
            "sample_index": col - data_start_col + 1,
            "sample_id": sample_id,
            "datetime": dt,
            "hour": hour_bucket(dt),
            "status_bits": status,
            "failure_bits": failure,
            "positions": {},
            "missing": {},
        }

        for metric in position_metrics:
            row = metrics[metric]
            pin, axis = pin_axis_from_name(metric)
            raw = row[col] if col < len(row) else ""
            value = parse_float(raw)
            record["positions"][(pin, axis)] = value
            record["missing"][(pin, axis)] = clean_cell(raw) == "" or clean_cell(raw).lower() == "nan"

        records.append(record)
    return records

# ============================================================
# ANALIZA
# ============================================================

def create_auto_pin_rows(active_pins):
    pins = sorted(active_pins)
    if PIN_ROWS_MODE == "CUSTOM" and CUSTOM_PIN_ROWS:
        return CUSTOM_PIN_ROWS
    if not pins:
        return {}

    max_pin = max(pins)
    rows = {}
    if max_pin <= 8:
        # typowy układ 2 kolumny x 4 rzędy
        for i in range(1, 5):
            group = [p for p in [i, i + 4] if p in pins]
            if len(group) >= 2:
                rows[f"ROW_{i}"] = group
    elif max_pin <= 16:
        # typowy układ 2 kolumny x 8 rzędów
        for i in range(1, 9):
            group = [p for p in [i, i + 8] if p in pins]
            if len(group) >= 2:
                rows[f"ROW_{i}"] = group
    else:
        # fallback: grupy po 4 kolejne piny
        for idx in range(0, len(pins), 4):
            group = pins[idx:idx + 4]
            if len(group) >= 2:
                rows[f"ROW_BLOCK_{idx//4 + 1}"] = group
    return rows


def is_fail_record(record):
    return (record.get("status_bits") not in [None, 0]) or (record.get("failure_bits") not in [None, 0])


def build_baseline(records):
    vals = defaultdict(list)
    quality = defaultdict(lambda: Counter())
    active_pins_by_group = defaultdict(set)

    for r in records:
        group = (r["file"], r["ref_name"])
        if ANALYZE_FAIL_ONLY and not is_fail_record(r):
            continue
        for (pin, axis), v in r["positions"].items():
            key = (r["file"], r["ref_name"], pin, axis)
            if r["missing"].get((pin, axis)):
                quality[key]["missing"] += 1
                continue
            if v is None:
                quality[key]["invalid"] += 1
                continue
            if v == 0:
                quality[key]["zero"] += 1
                if IGNORE_ZERO_VALUES:
                    continue
            if v < GROSS_MIN or v > GROSS_MAX:
                quality[key]["gross_for_baseline"] += 1
                continue
            vals[key].append(v)
            active_pins_by_group[group].add(pin)
            quality[key]["valid"] += 1

    baseline = {}
    for key, values in vals.items():
        if len(values) < MIN_VALID_VALUES_FOR_ANALYSIS:
            continue
        med, mad = robust_stats(values)
        baseline[key] = {
            "median": med,
            "mad": mad,
            "n": len(values),
            "min": min(values),
            "max": max(values),
            "mean": sum(values) / len(values),
        }
    return baseline, quality, active_pins_by_group


def classify(record, pin, axis, value, base):
    if value is None:
        return None
    if IGNORE_ZERO_VALUES and value == 0:
        return None
    if base is None:
        return None

    med = base["median"]
    mad = base["mad"]
    delta = value - med
    rz = robust_z(value, med, mad)
    fallback = FALLBACK_TOLERANCE.get(axis, 0.1)

    reason = None
    if value < GROSS_MIN or value > GROSS_MAX:
        reason = "GROSS_OUTLIER"
    elif rz is not None and abs(rz) >= ROBUST_Z_THRESHOLD:
        reason = f"ROBUST_Z>{ROBUST_Z_THRESHOLD}"
    elif (mad is None or mad == 0) and abs(delta) >= fallback:
        reason = f"ABS_DELTA>{fallback}"

    if not reason:
        return None

    return {
        "file": record["file"],
        "ref_name": record["ref_name"],
        "sample_index": record["sample_index"],
        "sample_id": record["sample_id"],
        "datetime": record["datetime"],
        "hour": record["hour"],
        "status_bits": record["status_bits"],
        "failure_bits": record["failure_bits"],
        "pin": pin,
        "axis": axis,
        "value": value,
        "median": med,
        "delta": delta,
        "abs_delta": abs(delta),
        "direction": "PLUS" if delta > 0 else "MINUS",
        "robust_z": rz,
        "reason": reason,
    }


def analyze(records):
    baseline, quality, active_pins_by_group = build_baseline(records)

    deviations = []
    row_events = []
    c_file = Counter()
    c_pin_axis = Counter()
    c_pin = Counter()
    c_axis = Counter()
    c_direction = Counter()
    c_hour = Counter()
    c_file_hour = Counter()
    c_row = Counter()
    c_row_axis = Counter()

    pin_rows_by_group = {g: create_auto_pin_rows(pins) for g, pins in active_pins_by_group.items()}

    for r in records:
        if ANALYZE_FAIL_ONLY and not is_fail_record(r):
            continue
        sample_devs = []
        for (pin, axis), value in r["positions"].items():
            key = (r["file"], r["ref_name"], pin, axis)
            dev = classify(r, pin, axis, value, baseline.get(key))
            if not dev:
                continue
            deviations.append(dev)
            sample_devs.append(dev)
            c_file[r["file"]] += 1
            c_pin_axis[(r["file"], r["ref_name"], pin, axis)] += 1
            c_pin[(r["file"], r["ref_name"], pin)] += 1
            c_axis[(r["file"], r["ref_name"], axis)] += 1
            c_direction[(r["file"], r["ref_name"], axis, dev["direction"])] += 1
            c_hour[r["hour"]] += 1
            c_file_hour[(r["file"], r["hour"])] += 1

        if sample_devs:
            group = (r["file"], r["ref_name"])
            rows_map = pin_rows_by_group.get(group, {})
            existing = {(d["pin"], d["axis"]): d for d in sample_devs}
            for row_name, pins in rows_map.items():
                for axis in ["X", "Y", "Z"]:
                    failed_pins = [p for p in pins if (p, axis) in existing]
                    if len(failed_pins) >= MIN_PINS_IN_ROW_FAIL:
                        event = {
                            "file": r["file"],
                            "ref_name": r["ref_name"],
                            "sample_index": r["sample_index"],
                            "sample_id": r["sample_id"],
                            "datetime": r["datetime"],
                            "hour": r["hour"],
                            "row_name": row_name,
                            "axis": axis,
                            "pins": ",".join(str(p) for p in failed_pins),
                            "count_pins": len(failed_pins),
                        }
                        row_events.append(event)
                        c_row[(r["file"], r["ref_name"], row_name)] += 1
                        c_row_axis[(r["file"], r["ref_name"], row_name, axis)] += 1

    return {
        "baseline": baseline,
        "quality": quality,
        "active_pins_by_group": active_pins_by_group,
        "pin_rows_by_group": pin_rows_by_group,
        "deviations": deviations,
        "row_events": row_events,
        "c_file": c_file,
        "c_pin_axis": c_pin_axis,
        "c_pin": c_pin,
        "c_axis": c_axis,
        "c_direction": c_direction,
        "c_hour": c_hour,
        "c_file_hour": c_file_hour,
        "c_row": c_row,
        "c_row_axis": c_row_axis,
    }

# ============================================================
# EXCEL
# ============================================================

def dt_text(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


def auto_width(ws, max_width=55):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = 10
        for cell in col:
            if cell.value is not None:
                width = max(width, min(max_width, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = width


def style_header(ws):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(bottom=thin)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def append_table(ws, headers, rows):
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_header(ws)
    auto_width(ws)


def add_bar_chart(ws, title, data_min_col, data_max_col, cat_col, start_row, end_row, pos):
    if end_row <= start_row:
        return
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "Kategoria"
    chart.x_axis.title = "Liczba"
    data = Reference(ws, min_col=data_min_col, max_col=data_max_col, min_row=start_row, max_row=end_row)
    cats = Reference(ws, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 15
    ws.add_chart(chart, pos)


def add_line_chart(ws, title, data_col, cat_col, start_row, end_row, pos):
    if end_row <= start_row:
        return
    chart = LineChart()
    chart.title = title
    chart.y_axis.title = "Liczba odchyleń"
    chart.x_axis.title = "Godzina"
    data = Reference(ws, min_col=data_col, min_row=start_row, max_row=end_row)
    cats = Reference(ws, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 18
    ws.add_chart(chart, pos)


def create_excel_report(output_xlsx, analysis, records, input_files):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    deviations = analysis["deviations"]
    row_events = analysis["row_events"]

    # Dashboard
    dash_rows = [
        ["Parametr", "Wartość"],
        ["Data raportu", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Liczba plików", len(input_files)],
        ["Liczba rekordów/pomiarów", len(records)],
        ["Liczba wykrytych odchyleń", len(deviations)],
        ["Liczba zdarzeń typu cały rząd", len(row_events)],
        ["Metoda", "Robust Z-score: mediana + MAD"],
        ["Próg robust Z", ROBUST_Z_THRESHOLD],
        ["Tryb tylko FAIL", str(ANALYZE_FAIL_ONLY)],
    ]
    for row in dash_rows:
        ws.append(row)
    style_header(ws)
    auto_width(ws)

    # Input files
    ws_files = wb.create_sheet("Input_Files")
    append_table(ws_files, ["Plik"], [[os.path.basename(f)] for f in input_files])

    # Summary per file
    summary_rows = []
    for file, cnt in analysis["c_file"].most_common():
        rec_count = sum(1 for r in records if r["file"] == file)
        summary_rows.append([file, rec_count, cnt])
    ws_sum = wb.create_sheet("Summary_File")
    append_table(ws_sum, ["File", "Records", "Deviation_Count"], summary_rows)
    add_bar_chart(ws_sum, "Odchylenia wg pliku", 3, 3, 1, 1, ws_sum.max_row, "E2")

    # Top pin axis
    rows = []
    for (file, ref, pin, axis), cnt in analysis["c_pin_axis"].most_common():
        rows.append([file, ref, f"Pin{pin}", axis, cnt])
    ws_tpa = wb.create_sheet("Top_Pin_Axis")
    append_table(ws_tpa, ["File", "Ref", "Pin", "Axis", "Deviation_Count"], rows)
    add_bar_chart(ws_tpa, "TOP Pin/Oś", 5, 5, 3, 1, min(ws_tpa.max_row, TOP_N + 1), "G2")

    # Top pins
    rows = []
    for (file, ref, pin), cnt in analysis["c_pin"].most_common():
        rows.append([file, ref, f"Pin{pin}", cnt])
    ws_tp = wb.create_sheet("Top_Pins")
    append_table(ws_tp, ["File", "Ref", "Pin", "Deviation_Count"], rows)
    add_bar_chart(ws_tp, "TOP Piny", 4, 4, 3, 1, min(ws_tp.max_row, TOP_N + 1), "F2")

    # Axis direction
    rows = []
    for (file, ref, axis, direction), cnt in analysis["c_direction"].most_common():
        rows.append([file, ref, axis, direction, cnt])
    ws_ad = wb.create_sheet("Axis_Direction")
    append_table(ws_ad, ["File", "Ref", "Axis", "Direction", "Deviation_Count"], rows)
    add_bar_chart(ws_ad, "Kierunek odchyleń wg osi", 5, 5, 3, 1, min(ws_ad.max_row, TOP_N + 1), "G2")

    # Trend hour
    rows = []
    for (file, h), cnt in sorted(analysis["c_file_hour"].items(), key=lambda x: (x[0][0], x[0][1])):
        rows.append([file, h, cnt])
    ws_tr = wb.create_sheet("Trend_Hour")
    append_table(ws_tr, ["File", "Hour", "Deviation_Count"], rows)
    add_line_chart(ws_tr, "Trend godzinowy odchyleń", 3, 2, 1, ws_tr.max_row, "E2")

    # Row problems
    rows = []
    for e in row_events:
        rows.append([e["file"], e["ref_name"], e["sample_index"], e["sample_id"], dt_text(e["datetime"]), e["hour"], e["row_name"], e["axis"], e["pins"], e["count_pins"]])
    ws_re = wb.create_sheet("Row_Problems")
    append_table(ws_re, ["File", "Ref", "Sample_Index", "Sample_ID", "Datetime", "Hour", "Row", "Axis", "Pins", "Count_Pins"], rows)

    # Baseline
    rows = []
    for (file, ref, pin, axis), st in sorted(analysis["baseline"].items()):
        rows.append([file, ref, f"Pin{pin}", axis, st["n"], st["median"], st["mad"], st["mean"], st["min"], st["max"]])
    ws_bl = wb.create_sheet("Baseline")
    append_table(ws_bl, ["File", "Ref", "Pin", "Axis", "N", "Median", "MAD", "Mean", "Min", "Max"], rows)

    # Data quality
    rows = []
    for (file, ref, pin, axis), q in sorted(analysis["quality"].items()):
        rows.append([file, ref, f"Pin{pin}", axis, q.get("valid", 0), q.get("zero", 0), q.get("missing", 0), q.get("invalid", 0), q.get("gross_for_baseline", 0)])
    ws_q = wb.create_sheet("Data_Quality")
    append_table(ws_q, ["File", "Ref", "Pin", "Axis", "Valid", "Zero", "Missing", "Invalid", "Gross_Excluded_From_Baseline"], rows)

    # Deviations
    rows = []
    for e in deviations[:MAX_DEVIATIONS_IN_EXCEL]:
        rows.append([
            e["file"], e["ref_name"], e["sample_index"], e["sample_id"], dt_text(e["datetime"]), e["hour"],
            e["status_bits"], e["failure_bits"], f"Pin{e['pin']}", e["axis"], e["value"], e["median"], e["delta"],
            e["abs_delta"], e["direction"], None if e["robust_z"] is None else round(e["robust_z"], 3), e["reason"]
        ])
    ws_dev = wb.create_sheet("Deviations")
    append_table(ws_dev, ["File", "Ref", "Sample_Index", "Sample_ID", "Datetime", "Hour", "StatusBits", "Failure_Bits", "Pin", "Axis", "Value", "Median", "Delta", "Abs_Delta", "Direction", "Robust_Z", "Reason"], rows)
    if ws_dev.max_row > 1:
        ws_dev.conditional_formatting.add(f"N2:N{ws_dev.max_row}", ColorScaleRule(start_type='min', start_color='FFFFFF', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))

    # Heatmap pin/axis counts
    heat_counts = analysis["c_pin_axis"]
    heat_rows = []
    for (file, ref, pin, axis), cnt in heat_counts.items():
        heat_rows.append([file, ref, pin, axis, cnt])
    ws_h = wb.create_sheet("Heatmap_Data")
    append_table(ws_h, ["File", "Ref", "Pin", "Axis", "Deviation_Count"], sorted(heat_rows))
    if ws_h.max_row > 1:
        ws_h.conditional_formatting.add(f"E2:E{ws_h.max_row}", ColorScaleRule(start_type='min', start_color='FFFFFF', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))

    # Pin rows config used
    rows = []
    for (file, ref), mapping in sorted(analysis["pin_rows_by_group"].items()):
        for row_name, pins in mapping.items():
            rows.append([file, ref, row_name, ",".join(str(p) for p in pins)])
    ws_cfg = wb.create_sheet("Pin_Rows_Config")
    append_table(ws_cfg, ["File", "Ref", "Row", "Pins"], rows)

    wb.save(output_xlsx)

# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) < 2:
        print("Przeciągnij jeden lub więcej plików CSV na program.")
        input("ENTER aby wyjść...")
        return

    csv_files = [p for p in sys.argv[1:] if p.lower().endswith(".csv") and os.path.isfile(p)]
    if not csv_files:
        print("Nie znaleziono plików CSV.")
        input("ENTER aby wyjść...")
        return

    all_records = []
    print("Analizowane pliki:")
    for file_path in csv_files:
        print("-", file_path)
        try:
            recs = extract_records(file_path)
            all_records.extend(recs)
            print(f"  OK: {len(recs)} rekordów")
        except Exception as e:
            print(f"  BŁĄD: {e}")

    if not all_records:
        print("Brak danych do analizy.")
        input("ENTER aby wyjść...")
        return

    analysis = analyze(all_records)

    print("\nPODSUMOWANIE")
    print("=" * 80)
    print(f"Rekordy: {len(all_records)}")
    print(f"Odchylenia: {len(analysis['deviations'])}")
    print(f"Problemy rzędów: {len(analysis['row_events'])}")
    print("\nTOP Pin/Oś:")
    for (file, ref, pin, axis), cnt in analysis["c_pin_axis"].most_common(TOP_N):
        print(f"{file} | {ref} | Pin{pin} {axis}: {cnt}")

    out_dir = get_exe_folder()
    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    output_xlsx = os.path.join(out_dir, f"pin_position_report__{timestamp}.xlsx")
    create_excel_report(output_xlsx, analysis, all_records, csv_files)

    print("\nZapisano raport Excel:")
    print(output_xlsx)
    input("\nENTER aby zamknąć...")


if __name__ == "__main__":
    main()
