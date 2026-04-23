import io
import re
import uuid
import os
import tempfile

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Section detection ────────────────────────────────────────────────────────

SECTION_PATTERNS = [
    ("vcd",        re.compile(r"viable cell density", re.I)),
    ("ivcd",       re.compile(r"\bivcd\b", re.I)),
    ("viability",  re.compile(r"\bviabilit", re.I)),
    ("titer",      re.compile(r"\btiter\b", re.I)),
    ("qp",         re.compile(r"\bqp\b", re.I)),
    ("mu",         re.compile(r"specific growth|[μµ]|\bmu\b", re.I)),
]

SECTION_META = {
    "vcd":        {"label": "Viable Cell Density (VCD)", "unit": "×10⁵ cells/mL",    "chart_type": "line"},
    "ivcd":       {"label": "IVCD",                      "unit": "×10⁵ cell·day/mL", "chart_type": "line"},
    "viability":  {"label": "Viability",                 "unit": "%",                "chart_type": "line"},
    "titer":      {"label": "Titer",                     "unit": "mg/L",             "chart_type": "line"},
    "qp":         {"label": "Specific Productivity (Qp)","unit": "pg/cell/day",      "chart_type": "bar"},
    "mu":         {"label": "Specific Growth Rate (μ)",  "unit": "day⁻¹",            "chart_type": "bar"},
}

FEEDING_DAYS = {
    "CHO":       {"Fed-batch": [0, 3, 6], "Batch": []},
    "Hybridoma": {"Fed-batch": [0, 2, 4], "Batch": []},
    "VERO":      {"Fed-batch": [],        "Batch": []},
}

# ── Helpers ──────────────────────────────────────────────────────────────────

def _detect_cell_line(title: str) -> str:
    t = title.upper()
    if "CHO" in t:        return "CHO"
    if "HYBRIDOMA" in t:  return "Hybridoma"
    if "VERO" in t:       return "VERO"
    return "Unknown"


def _detect_culture_mode(title: str) -> str:
    t = title.upper()
    if "FED" in t:    return "Fed-batch"
    if "BATCH" in t:  return "Batch"
    return "Unknown"


def _find_정리_sheet(xl: pd.ExcelFile) -> str:
    for name in xl.sheet_names:
        if "정리" in name or "결과" in name:
            return name
    return xl.sheet_names[-1]


def _find_section_boundaries(df: pd.DataFrame) -> dict:
    found = {}
    for r in range(len(df)):
        cell = str(df.iloc[r, 0]) if pd.notna(df.iloc[r, 0]) else ""
        for sec, pat in SECTION_PATTERNS:
            if sec not in found and pat.search(cell):
                found[sec] = r
                break
    return found


def _find_day_row(df: pd.DataFrame, start: int, end: int):
    """Return (row_idx, [days], [col_indices]) for the row with 0,1,2,... integers."""
    for r in range(start, min(end, start + 12)):
        if r >= len(df):
            break
        row = df.iloc[r]
        candidates = [
            (c, int(v))
            for c, v in enumerate(row)
            if isinstance(v, (int, float))
            and not pd.isna(v)
            and v == int(v)
            and int(v) >= 0
        ]
        if len(candidates) < 3:
            continue
        vals = [v for _, v in candidates]
        # Must start from 0 and be consecutive
        if vals[0] == 0 and vals == list(range(len(vals))):
            return r, vals, [c for c, _ in candidates]
    return None, [], []


def _parse_section(df: pd.DataFrame, start: int, end: int):
    """Return (days, {treatment: [[rep1_vals], [rep2_vals], ...]})."""
    day_row, days, day_cols = _find_day_row(df, start + 1, end)
    if day_row is None:
        return [], {}

    treatments: dict[str, list] = {}
    current = None

    for r in range(day_row + 1, min(end, len(df))):
        row = df.iloc[r]
        cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""

        if cell and cell.lower() not in ("nan", ""):
            current = cell

        if current is None:
            continue

        vals = []
        for c in day_cols:
            v = row.iloc[c] if c < len(row) else None
            if v is None or pd.isna(v) or str(v).strip() in ("-", "", "nan"):
                vals.append(None)
            else:
                try:
                    vals.append(float(v))
                except (ValueError, TypeError):
                    vals.append(None)

        if any(v is not None for v in vals):
            treatments.setdefault(current, []).append(vals)

    return days, treatments


def _compute_stats(treatments: dict) -> dict:
    stats = {}
    for name, reps in treatments.items():
        n = max(len(r) for r in reps)
        means, stds = [], []
        for d in range(n):
            vals = [r[d] for r in reps if d < len(r) and r[d] is not None]
            if vals:
                means.append(round(float(np.mean(vals)), 4))
                stds.append(round(float(np.std(vals, ddof=1)) if len(vals) > 1 else 0.0, 4))
            else:
                means.append(None)
                stds.append(None)
        stats[name] = {"mean": means, "std": stds, "replicates": reps}
    return stats

def _fix_interval_labels(sections: dict):
    """Replace 'D6' style bar-chart labels with 'D3-D6' interval style."""
    # μ: intervals between consecutive non-None VCD measurement days
    if "mu" in sections:
        mu_days = sections["mu"]["days"]
        prev = 0
        labels = []
        for d in mu_days:
            labels.append(f"D{prev}-D{d}")
            prev = d
        sections["mu"]["x_labels"] = labels

    # Qp: intervals between consecutive Titer measurement days
    if "qp" in sections:
        qp_days = sections["qp"]["days"]
        # Use Titer non-None days as reference for interval starts
        titer_non_none = []
        if "titer" in sections:
            t_days = sections["titer"]["days"]
            first_means = next(iter(sections["titer"]["treatments"].values()), {}).get("mean", [])
            titer_non_none = [d for d, m in zip(t_days, first_means) if m is not None]
        labels = []
        for qd in qp_days:
            prev = max([t for t in titer_non_none if t < qd], default=0)
            labels.append(f"D{prev}-D{qd}")
        sections["qp"]["x_labels"] = labels


# ── Public API ───────────────────────────────────────────────────────────────

def process_file(file_bytes: bytes) -> dict:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = _find_정리_sheet(xl)
    df = xl.parse(sheet_name, header=None)

    title = str(df.iloc[0, 0]) if pd.notna(df.iloc[0, 0]) else ""
    cell_line = _detect_cell_line(title)
    culture_mode = _detect_culture_mode(title)

    boundaries = _find_section_boundaries(df)
    sorted_secs = sorted(boundaries.items(), key=lambda x: x[1])

    result_sections = {}
    for i, (sec, start) in enumerate(sorted_secs):
        end = sorted_secs[i + 1][1] if i + 1 < len(sorted_secs) else len(df)
        days, treatments = _parse_section(df, start, end)
        if not treatments:
            continue

        stats = _compute_stats(treatments)

        # For bar charts keep only days that have at least one non-None value
        if SECTION_META.get(sec, {}).get("chart_type") == "bar":
            valid_idx = [
                i for i, d in enumerate(days)
                if any(
                    t["mean"][i] is not None
                    for t in stats.values()
                    if i < len(t["mean"])
                )
            ]
            days = [days[i] for i in valid_idx]
            for t in stats.values():
                t["mean"] = [t["mean"][i] for i in valid_idx if i < len(t["mean"])]
                t["std"]  = [t["std"][i]  for i in valid_idx if i < len(t["std"])]
                t["replicates"] = [
                    [rep[i] for i in valid_idx if i < len(rep)]
                    for rep in t["replicates"]
                ]

        result_sections[sec] = {
            **SECTION_META.get(sec, {"label": sec, "unit": "", "chart_type": "line"}),
            "days": days,
            "x_labels": [f"D{d}" for d in days],
            "treatments": stats,
        }

    # Fix interval labels for bar chart sections
    _fix_interval_labels(result_sections)

    feeding_days = FEEDING_DAYS.get(cell_line, {}).get(culture_mode, [])

    # Build and save Excel
    file_id = str(uuid.uuid4())
    out_path = os.path.join(tempfile.gettempdir(), f"{file_id}.xlsx")
    _build_excel(result_sections, cell_line, culture_mode, title, out_path)

    return {
        "cell_line": cell_line,
        "culture_mode": culture_mode,
        "feeding_days": feeding_days,
        "sections": result_sections,
        "title": title,
        "file_id": file_id,
    }

# ── Excel output ─────────────────────────────────────────────────────────────

_HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
_SUBHDR_FILL  = PatternFill("solid", fgColor="2E75B6")
_ALT_FILL     = PatternFill("solid", fgColor="D6E4F0")
_WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")
_THIN = Side(style="thin", color="AAAAAA")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _hdr_cell(ws, row, col, value, fill=None, font_color="FFFFFF", bold=True, size=11):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, color=font_color, name="Arial", size=size)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = _BORDER
    if fill:
        c.fill = fill
    return c


def _data_cell(ws, row, col, value, fill=None, number_format="0.000"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", size=10)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = _BORDER
    c.number_format = number_format
    if fill:
        c.fill = fill
    return c


def _build_excel(sections: dict, cell_line: str, culture_mode: str, title: str, path: str):
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for sec_key, sec in sections.items():
        ws = wb.create_sheet(title=sec["label"][:31])
        days = sec["days"]
        treatments = sec["treatments"]
        chart_type = sec["chart_type"]

        # Title row
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + len(days))
        _hdr_cell(ws, 1, 1, f"{title}  |  {sec['label']} ({sec['unit']})",
                  fill=_HEADER_FILL, size=12)
        ws.row_dimensions[1].height = 22

        # Column headers
        x_labels = sec["x_labels"]
        _hdr_cell(ws, 2, 1, "Treatment", fill=_SUBHDR_FILL)
        _hdr_cell(ws, 2, 2, "Stat",      fill=_SUBHDR_FILL)
        for j, lbl in enumerate(x_labels):
            _hdr_cell(ws, 2, 3 + j, lbl, fill=_SUBHDR_FILL)
        ws.row_dimensions[2].height = 18

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 8
        for j in range(len(days)):
            ws.column_dimensions[get_column_letter(3 + j)].width = 10

        # Data rows
        row = 3
        for t_idx, (name, stat) in enumerate(treatments.items()):
            fill = _ALT_FILL if t_idx % 2 == 0 else _WHITE_FILL

            # Mean row
            ws.merge_cells(start_row=row, start_column=1, end_row=row + 1, end_column=1)
            name_cell = ws.cell(row=row, column=1, value=name)
            name_cell.font = Font(bold=True, name="Arial", size=10)
            name_cell.alignment = Alignment(horizontal="left", vertical="center")
            name_cell.border = _BORDER
            name_cell.fill = fill

            _hdr_cell(ws, row, 2, "Mean", fill=fill, font_color="000000", bold=False)
            for j, v in enumerate(stat["mean"]):
                _data_cell(ws, row, 3 + j, v, fill=fill)

            # SD row
            _hdr_cell(ws, row + 1, 2, "SD", fill=fill, font_color="000000", bold=False)
            for j, v in enumerate(stat["std"]):
                _data_cell(ws, row + 1, 3 + j, v, fill=fill)

            row += 2

        # Individual replicates (collapsed below)
        sep_row = row
        ws.merge_cells(start_row=sep_row, start_column=1, end_row=sep_row, end_column=3 + len(days))
        _hdr_cell(ws, sep_row, 1, "── Individual Replicates ──",
                  fill=_SUBHDR_FILL, size=10)
        row = sep_row + 1

        for t_idx, (name, stat) in enumerate(treatments.items()):
            for rep_i, rep in enumerate(stat["replicates"]):
                fill = _ALT_FILL if t_idx % 2 == 0 else _WHITE_FILL
                lbl = f"{name} (rep {rep_i + 1})"
                c = ws.cell(row=row, column=1, value=lbl)
                c.font = Font(name="Arial", size=9, italic=True)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = _BORDER
                c.fill = fill
                ws.cell(row=row, column=2, value="").border = _BORDER
                for j, v in enumerate(rep[:len(days)]):
                    _data_cell(ws, row, 3 + j, v, fill=fill)
                row += 1

    wb.save(path)
