from __future__ import annotations

import os
import io
import math
from datetime import datetime, date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

EXCEL_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "DTMC stats.xlsx"
)

APP_VERSION = "v20 — 2026-07-13"

REPORT_TITLE = "Financial Performance of RG Participants - FY 2024 and FY 2025"

METRIC_COL = "In MM (USD)"

SCALE_NOTE = "Revenue and Net Income are in USD millions ($MM)."

CONSOLIDATED_NOTE = (
    "Merrill Lynch and Citibank NA reflect the consolidated financials of "
    "Bank of America and Citigroup, respectively."
)

# Region for each bank column. Banks not listed fall back to "Other" — add new
# banks here to place them in a region.
BANK_REGIONS = {
    "TD": "Canada",
    "RBC": "Canada",
    "BNS": "Canada",
    "BMO": "Canada",
    "CIBC": "Canada",
    "NBC": "Canada",
    "LBC": "Canada",
    "Desjardins": "Canada",
    "ATB": "Canada",
    "BNP Paribas": "Europe",
    "Merrill Lynch": "United States",
    "Citibank NA": "United States",
}


def bank_region(bank):
    return BANK_REGIONS.get(str(bank).strip(), "Other")


# Full legal/common names for the legend at the bottom of the page and in
# the PDF. Banks not listed fall back to their column name.
BANK_FULL_NAMES = {
    "TD": "Toronto-Dominion Bank",
    "RBC": "Royal Bank of Canada",
    "BNS": "Bank of Nova Scotia (Scotiabank)",
    "BMO": "Bank of Montreal",
    "CIBC": "Canadian Imperial Bank of Commerce",
    "NBC": "National Bank of Canada",
    "LBC": "Laurentian Bank of Canada",
    "Desjardins": "Desjardins Group",
    "ATB": "ATB Financial",
    "BNP Paribas": "BNP Paribas S.A.",
    "Merrill Lynch": "Merrill Lynch (Bank of America)",
    "Citibank NA": "Citibank N.A. (Citigroup)",
}


def full_bank_name(bank):
    return BANK_FULL_NAMES.get(str(bank).strip(), str(bank))


# Reporting-date overrides: these take precedence over the Date row in the
# Excel file. Remove an entry (or the whole dict) to fall back to the sheet.
REPORTING_DATE_OVERRIDES = {
    "TD": "October-2025",
    "RBC": "October-2025",
    "BNS": "October-2025",
    "BMO": "October-2025",
    "CIBC": "October-2025",
    "NBC": "October-2025",
    "LBC": "October-2025",
    "Desjardins": "December-2025",
    "BNP Paribas": "December-2025",
    "Merrill Lynch": "December-2025",
    "Citibank NA": "December-2025",
    "ATB": "March-2025",
}


# Short display names for chart x-axes (full name still shown on hover).
# Banks not listed fall back to their full name.
SHORT_NAMES = {
    "BNP Paribas": "BNP",
    "Merrill Lynch": "ML",
    "Citibank NA": "Citi",
    "Desjardins": "DESJ",
}


def short_name(bank):
    return SHORT_NAMES.get(str(bank).strip(), str(bank))


def compact_label(value, fmt):
    """On-bar labels. Currency shows the full $MM value; percents and plain
    numbers drop trailing zeros (142% not 142.00%, 4 not 4.00) so labels stay
    narrow and don't collide."""
    if fmt == "currency":
        return f"${value:,.0f}"
    if fmt == "percent":
        return f"{value:,.1f}%"
    return f"{value:,.1f}"


MISSING_TOKENS = {"", "na", "n/a", "n.a.", "-", "—", "nm", "nmf"}

# Rows with these names (case-insensitive) are treated as per-bank reporting
# dates, not metrics: pulled out of the charts/heatmap and shown as captions.
DATE_ROW_NAMES = {"date", "as of", "as-of", "as of date", "reporting date"}

PERFORMANCE_KEYWORDS = (
    "revenue", "income", "earnings", "profit", "margin",
    "yoy", "growth", "eps", "roe", "roa"
)

# Checked BEFORE the performance keywords, so "Equity Price" doesn't get
# claimed by another topic.
MARKET_KEYWORDS = (
    "cds", "equity price", "share price", "stock", "spread", "market"
)

TOPIC_PERFORMANCE = "Financial Performance"
TOPIC_RISK = "Capital, Liquidity & Asset Quality"
TOPIC_MARKET = "Market Indicators"

TOPIC_ICONS = {
    TOPIC_PERFORMANCE: "📈",
    TOPIC_RISK: "🛡️",
    TOPIC_MARKET: "💹",
}

# CDS charts invert the color logic: a POSITIVE change (spread widening) is
# BAD news for credit perception, a NEGATIVE change (tightening) is good.
MARKET_NOTE = (
    "Market indicators are as of June 2026 — the ▲ 3 months change is "
    "measured vs. March 2026, and the ▲ 1 year change vs. June 2025."
)

# (internal metric name, bank) pairs excluded from charts and the heatmap.
# "YoY (▲%) (2)" is the Net Income YoY row (second YoY row in the sheet).
CHART_EXCLUSIONS = {("YoY (▲%) (2)", "LBC")}


def chart_excluded_banks(metric):
    return [b for (m, b) in CHART_EXCLUSIONS if m == metric]


ACCENT = "#1f6f8b"   # neutral bars
NEG = "#c0392b"      # negative values only (positive for CDS widening)
POS = "#2a9d4a"      # CDS tightening only
HILITE = "#e08a1e"   # highlighted bank
GRID_CLR = "#eef2f4"
ZERO_CLR = "#c9d2d8"
STRIPE_BG = "#f8fbfc"


def make_unique_index(index):
    counts = {}
    new_index = []

    for item in index:
        item = str(item).strip()

        if item not in counts:
            counts[item] = 1
            new_index.append(item)
        else:
            counts[item] += 1
            new_index.append(f"{item} ({counts[item]})")

    return new_index


def clean_metric_name(metric):
    metric = str(metric)

    if metric.endswith(")") and " (" in metric:
        base, suffix = metric.rsplit(" (", 1)
        if suffix[:-1].isdigit():
            return base

    return metric


def is_cds_metric(metric):
    return "cds" in clean_metric_name(metric).lower()


def display_metric_name(metric):
    """Chart/table title: cleaned name with '%' markers removed —
    'Equity Price %(▲ 3 months)' → 'Equity Price (▲ 3 months)',
    'Gross NPAs/... + OREO (%)' → 'Gross NPAs/... + OREO',
    'YoY (▲%)' → 'YoY (▲)'."""
    s = clean_metric_name(metric)
    s = s.replace("(▲%)", "(▲)")
    s = s.replace("%(", "(")
    s = s.replace("(%)", "")
    return " ".join(s.split())


YEAR_COLORS = {
    "2024": "#8a9aa5",
    "2025": ACCENT,
}


def metric_year(metric):
    """Return a trailing reporting year such as '2024' or '2025'."""
    import re
    match = re.search(r"\((20\d{2})\)\s*$", clean_metric_name(metric))
    return match.group(1) if match else None


def metric_base_name(metric):
    """Remove a trailing year from a metric for combined-year chart titles."""
    import re
    name = clean_metric_name(metric)
    name = re.sub(r"\s*\((20\d{2})\)\s*$", "", name)
    return " ".join(name.split())


def build_chart_groups(metrics, selected_years=("2024", "2025")):
    """Build one chart definition per metric.

    Annual rows such as Revenue (2024) and Revenue (2025) are combined into
    one grouped chart, but only for the years selected by the user. Rows
    without a year, including YoY and market indicators, remain individual
    charts.
    """
    groups = []
    annual_positions = {}

    for metric in metrics:
        year = metric_year(metric)
        if year in {"2024", "2025"}:
            if year not in set(selected_years):
                continue
            base = metric_base_name(metric)
            if base not in annual_positions:
                annual_positions[base] = len(groups)
                groups.append({"title": base, "metrics": {}})
            groups[annual_positions[base]]["metrics"][year] = metric
        elif year:
            # Keep other dated rows separate rather than mixing them into the
            # FY2024/FY2025 comparison.
            groups.append({
                "title": display_metric_name(metric),
                "metrics": {year: metric},
            })
        else:
            groups.append({
                "title": display_metric_name(metric),
                "metrics": {"": metric},
            })

    for group in groups:
        group["metrics"] = dict(sorted(group["metrics"].items()))

    return [g for g in groups if g["metrics"]]

def display_labels(metrics):
    """Cleaned metric names made unique again for pandas (Styler requires a
    unique index). Duplicates get zero-width spaces appended, so 'YoY (▲%)'
    can appear twice looking identical while remaining distinct labels."""
    seen = {}
    labels = []
    for m in metrics:
        name = display_metric_name(m)
        seen[name] = seen.get(name, 0) + 1
        labels.append(name + "\u200b" * (seen[name] - 1))
    return labels


def parse_value(raw):
    if raw is None:
        return None

    try:
        if pd.isna(raw):
            return None
    except Exception:
        pass

    s = str(raw).strip()

    if s.lower() in MISSING_TOKENS:
        return None

    negative = False

    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    s = s.replace("$", "").replace(",", "").replace("%", "").strip()

    if s.startswith("-"):
        negative = True
        s = s[1:]

    try:
        value = float(s)
    except ValueError:
        return None

    return -value if negative else value


def detect_format(metric_name, raw_values):
    name = clean_metric_name(metric_name).lower()
    cells = [str(v) for v in raw_values if v is not None]

    currency_keywords = ["revenue", "income"]
    percent_keywords = [
        "%", "yoy", "ratio", "lcr", "npas", "loans",
        "equity price", "cet", "roe", "roa"
    ]

    if any(k in name for k in currency_keywords):
        return "currency"

    if any(k in name for k in percent_keywords) or any("%" in c for c in cells):
        return "percent"

    if any("$" in c for c in cells):
        return "currency"

    return "number"


def format_value(value, fmt, original=""):
    try:
        if pd.isna(value):
            return str(original) if original not in (None, "") else "—"
    except Exception:
        return str(original)

    try:
        value = float(value)
    except Exception:
        return str(original)

    if fmt == "currency":
        return f"${value:,.0f}"

    if fmt == "percent":
        # Values arrive already scaled (the loader converts Excel's stored
        # fractions using each cell's display format), so no ×100 guessing.
        return f"{value:,.1f}%"

    return f"{value:,.1f}"


def group_metrics_by_topic(metrics, formats):
    perf, risk, market = [], [], []

    for m in metrics:
        name = clean_metric_name(m).lower()

        if any(k in name for k in MARKET_KEYWORDS):
            market.append(m)
        elif formats.get(m) == "currency" or any(k in name for k in PERFORMANCE_KEYWORDS):
            perf.append(m)
        else:
            risk.append(m)

    # Fallback: if everything landed in a single bucket, split evenly across
    # the first two topics so the pages are still useful.
    buckets = [perf, risk, market]
    if metrics and sum(1 for b in buckets if b) == 1:
        half = math.ceil(len(metrics) / 2)
        perf, risk, market = list(metrics[:half]), list(metrics[half:]), []

    return [
        (TOPIC_PERFORMANCE, perf),
        (TOPIC_RISK, risk),
        (TOPIC_MARKET, market),
    ]


def load_raw(source):
    """Load the Excel sheet using each cell's DISPLAY format, so a cell stored
    as 1.42 with a percent format arrives as '142%', 0.0054 as '0.54%', and
    59180 with a '$' format as '$59,180'. This fixes the LCR scaling issue and
    removes the need for any ×100 heuristics downstream. Dates are rendered
    per their format ('mmm-yy' → 'Apr-26', 'dd-mmm-yy' → '01-Mar-26'). Text
    like 'Meets Req' or 'NA' passes through verbatim."""
    from openpyxl import load_workbook

    wb = load_workbook(source, data_only=True)
    ws = wb.active

    def cell_to_str(cell):
        v = cell.value

        if v is None:
            return ""

        if isinstance(v, str):
            return v.strip()

        fmt = (cell.number_format or "").lower()

        if isinstance(v, (datetime, date)):
            if "d" in fmt:
                return v.strftime("%d-%b-%y")
            return v.strftime("%b-%y")

        if "%" in fmt:
            return f"{round(v * 100, 4):g}%"

        if "$" in fmt:
            return f"${v:,.0f}"

        if isinstance(v, float) and v == int(v):
            return str(int(v))

        return str(v)

    rows = [[cell_to_str(c) for c in row] for row in ws.iter_rows()]

    if not rows:
        return pd.DataFrame(columns=[METRIC_COL]).set_index(METRIC_COL)

    header = [h if h else f"Column {i}" for i, h in enumerate(rows[0])]
    header[0] = METRIC_COL

    body = [r for r in rows[1:] if r and str(r[0]).strip()]

    df = pd.DataFrame(body, columns=header)
    df = df.set_index(METRIC_COL)
    df.index = make_unique_index(df.index)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith("Column ")]]

    return df


def split_date_row(raw):
    """Pull the per-bank reporting-date row (e.g. 'Date') out of the metrics.

    Returns (metrics_frame, date_series_or_None)."""
    for m in raw.index:
        if clean_metric_name(m).strip().lower() in DATE_ROW_NAMES:
            dates = raw.loc[m]
            return raw.drop(index=m), dates
    return raw, None


def long_date_label(date_str):
    """'Apr-26' or '01-Mar-26' → 'April-2026'. Unparseable text passes
    through as-is."""
    s = str(date_str).strip()

    if not s or s == "—":
        return "—"

    d = None
    for fmt in ("%b-%y", "%d-%b-%y"):
        try:
            d = pd.to_datetime(s, format=fmt)
            break
        except Exception:
            continue

    if d is None:
        try:
            d = pd.to_datetime(s, dayfirst=True)
        except Exception:
            return s

    return f"{d.strftime('%B')}-{d.year}"


def date_caption_lines(dates, banks):
    """One line per reporting date, grouping the banks that share it:
    ['April-2026 — TD, RBC, …', 'December-2025 — Desjardins', …]"""
    if dates is None:
        return []

    groups = {}
    for b in banks:
        d = long_date_label(dates.get(b, ""))
        groups.setdefault(d, []).append(b)

    return [f"{d} — {', '.join(bs)}" for d, bs in groups.items()]


def build_numeric(raw):
    formats = {}
    numeric_rows = []

    for metric, row in raw.iterrows():
        formats[metric] = detect_format(metric, row.tolist())
        numeric_rows.append([parse_value(v) for v in row])

    numeric = pd.DataFrame(
        numeric_rows,
        index=raw.index,
        columns=raw.columns
    )

    return numeric.astype("float64"), formats


# --------------------------------------------------------------------------- #
# PDF report (reportlab for layout, matplotlib for static charts)
# --------------------------------------------------------------------------- #
def _metric_png(chart_group, num_v, formats, sort_mode, highlight):
    """Render one grouped FY2024/FY2025 chart to PNG bytes."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    import numpy as np

    metric_items = list(chart_group["metrics"].items())
    if not metric_items:
        return None

    first_metric = metric_items[0][1]
    fmt = formats[first_metric]

    available = []
    for bank in num_v.columns:
        if any(pd.notna(num_v.loc[m, bank]) for _, m in metric_items):
            available.append(bank)

    available = [
        bank for bank in available
        if not all(bank in chart_excluded_banks(m) for _, m in metric_items)
    ]
    if not available:
        return None

    _, reference_metric = metric_items[-1]
    reference = num_v.loc[reference_metric, available]

    if sort_mode == "By value":
        order = reference.sort_values(
            ascending=False, na_position="last"
        ).index.tolist()
    elif sort_mode == "By region, then value":
        order = sorted(
            available,
            key=lambda b: (
                bank_region(b),
                -(reference.get(b) if pd.notna(reference.get(b))
                  else float("-inf")),
            ),
        )
    else:
        order = available

    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=150)
    x = np.arange(len(order))
    multiple_years = len(metric_items) > 1
    width = 0.36 if multiple_years else 0.62

    all_bars = []
    for i, (year_label, metric) in enumerate(metric_items):
        values = num_v.loc[metric, order]
        numeric_values = [float(v) if pd.notna(v) else np.nan for v in values]

        if multiple_years:
            offsets = (i - (len(metric_items) - 1) / 2) * width
            positions = x + offsets
            colors = [YEAR_COLORS.get(year_label, ACCENT)] * len(order)
        else:
            positions = x
            if is_cds_metric(metric):
                colors = [
                    NEG if pd.notna(v) and v > 0 else
                    POS if pd.notna(v) and v < 0 else ACCENT
                    for v in numeric_values
                ]
            else:
                colors = [
                    NEG if pd.notna(v) and v < 0 else ACCENT
                    for v in numeric_values
                ]

        if highlight != "(none)":
            colors = [
                HILITE if bank == highlight else color
                for bank, color in zip(order, colors)
            ]

        bars = ax.bar(
            positions,
            numeric_values,
            width=width,
            color=colors,
            label=f"FY{year_label}" if year_label else chart_group["title"],
        )
        all_bars.append((bars, numeric_values))

    ax.grid(axis="y", color=GRID_CLR, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.margins(y=0.20)
    ax.set_title(display_metric_name(chart_group["title"]),
                 fontsize=10, fontweight="bold")
    ax.axhline(0, color="#888888", linewidth=0.6)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_xticks(x)
    ax.set_xticklabels([short_name(b) for b in order], rotation=45,
                       ha="right", fontsize=7)
    ax.tick_params(axis="y", labelsize=7)

    if fmt == "currency":
        ax.set_ylabel("US$ MM", fontsize=7, color="#666666")
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
    elif fmt == "percent":
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.1f}%"))
    else:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.1f}"))

    for bars, vals in all_bars:
        labels = [compact_label(v, fmt) if pd.notna(v) else "" for v in vals]
        ax.bar_label(bars, labels=labels, fontsize=5.5, padding=2)

    if any(bool(year_label) for year_label, _ in metric_items):
        ax.legend(fontsize=7, frameon=False, ncol=min(2, len(metric_items)),
                  loc="upper right")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()

def build_pdf_report(raw_v, num_v, formats, source_label,
                     sort_mode, highlight, dates_items, selected_years):
    """Assemble a landscape PDF: title + comparison table (with 'Region' and
    'As of' columns) on page 1, then charts grouped by topic -- each topic
    starts on its own landscape page."""
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, Image, PageBreak)

    metrics = list(raw_v.index)
    banks = list(raw_v.columns)
    dates = dict(dates_items) if dates_items else {}
    page_w, page_h = landscape(letter)
    margin = 0.5 * inch
    avail_w = page_w - 2 * margin

    styles = getSampleStyleSheet()
    h_cell = ParagraphStyle("hcell", parent=styles["Normal"], fontSize=6.5,
                            leading=8, textColor=colors.white,
                            fontName="Helvetica-Bold")
    row_lbl = ParagraphStyle("rowlbl", parent=styles["Normal"], fontSize=7,
                             leading=8, fontName="Helvetica-Bold")
    subtitle = ParagraphStyle("sub", parent=styles["Normal"], fontSize=8,
                              textColor=colors.HexColor("#666666"))

    story = [
        Paragraph(REPORT_TITLE, styles["Title"]),
        Paragraph(
            f"Generated {datetime.now():%Y-%m-%d %H:%M} &nbsp;|&nbsp; "
            f"{len(banks)} banks &nbsp;|&nbsp; {len(metrics)} metrics "
            f"&nbsp;|&nbsp; source: {source_label} &nbsp;|&nbsp; {APP_VERSION}",
            subtitle,
        ),
        Spacer(1, 10),
    ]

    # --- Comparison table (banks as rows so it grows down the page).
    # Market indicators are charted on their own page but excluded here.
    table_metrics = [
        m for m in metrics
        if not any(k in clean_metric_name(m).lower() for k in MARKET_KEYWORDS)
    ]

    header = [Paragraph("Bank", h_cell)]
    header += [Paragraph(display_metric_name(m), h_cell)
               for m in table_metrics]

    table_data = [header]
    red_cells = []  # (col, row) coords of negative values

    for r, bank in enumerate(banks, start=1):
        cells = [Paragraph(bank, row_lbl)]
        for c, metric in enumerate(table_metrics, start=1):
            val = num_v.loc[metric, bank]
            cells.append(format_value(val, formats[metric], raw_v.loc[metric, bank]))
            if pd.notna(val) and val < 0:
                red_cells.append((c, r))
        table_data.append(cells)

    first_w = 1.0 * inch
    other_w = (avail_w - first_w) / max(len(table_metrics), 1)
    col_widths = [first_w] + [other_w] * len(table_metrics)
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(ACCENT)),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f3f7f9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for c, r in red_cells:
        style.append(("TEXTCOLOR", (c, r), (c, r), colors.HexColor(NEG)))
    table.setStyle(TableStyle(style))
    story.append(table)

    # Bank legend under the table.
    legend = ";  ".join(f"<b>{b}</b> = {full_bank_name(b)}" for b in banks)
    story += [
        Spacer(1, 6),
        Paragraph(legend, subtitle),
        Spacer(1, 4),
        Paragraph(SCALE_NOTE, subtitle),
        Paragraph(CONSOLIDATED_NOTE, subtitle),
    ]

    # --- Charts grouped by topic: each topic starts on its own page.
    img_w = (avail_w - 0.2 * inch) / 2
    img_h = img_w * 0.54

    chart_groups = build_chart_groups(metrics, selected_years)
    topic_groups = {
        TOPIC_PERFORMANCE: [],
        TOPIC_RISK: [],
        TOPIC_MARKET: [],
    }

    for group in chart_groups:
        representative = next(iter(group["metrics"].values()))
        name = clean_metric_name(representative).lower()
        if any(k in name for k in MARKET_KEYWORDS):
            topic = TOPIC_MARKET
        elif (formats.get(representative) == "currency"
              or any(k in name for k in PERFORMANCE_KEYWORDS)):
            topic = TOPIC_PERFORMANCE
        else:
            topic = TOPIC_RISK
        topic_groups[topic].append(group)

    for topic, page_groups in topic_groups.items():
        if not page_groups:
            continue

        story += [PageBreak(), Paragraph(topic, styles["Heading2"])]

        if topic == TOPIC_MARKET:
            story.append(Paragraph(MARKET_NOTE, subtitle))
        else:
            for line in date_caption_lines(
                    pd.Series(dates) if dates else None, banks):
                story.append(Paragraph(line, subtitle))
            year_text = " and ".join(f"FY{y}" for y in selected_years)
            if len(selected_years) > 1:
                year_text += " are displayed side by side."
            else:
                year_text += " is displayed."
            story.append(Paragraph(year_text, subtitle))
        story.append(Spacer(1, 4))

        pair = []
        chart_rows = []
        for group in page_groups:
            png = _metric_png(group, num_v, formats, sort_mode, highlight)
            if png is None:
                continue
            pair.append(Image(io.BytesIO(png), width=img_w, height=img_h))
            if len(pair) == 2:
                chart_rows.append(pair)
                pair = []
        if pair:
            pair.append("")
            chart_rows.append(pair)
        if chart_rows:
            grid = Table(chart_rows, colWidths=[img_w + 0.1 * inch] * 2)
            grid.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))
            story.append(grid)

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#999999"))
        canvas.drawRightString(page_w - margin, 0.3 * inch, f"Page {doc.page}")
        canvas.drawString(margin, 0.3 * inch,
                          f"{REPORT_TITLE} — Revenue & Net Income in MM (USD)")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(letter),
                            leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=0.6 * inch,
                            title=REPORT_TITLE)
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


@st.cache_data(show_spinner="Building PDF report…")
def get_report_bytes(raw_csv, formats_items, metrics, banks, source_label,
                     sort_mode, highlight, dates_items, selected_years,
                     app_version=APP_VERSION):
    """Cached wrapper so the PDF only rebuilds when the data, selection,
    sort order, or highlight changes."""
    raw_v = pd.read_csv(io.StringIO(raw_csv), dtype=str,
                        keep_default_na=False).fillna("")
    raw_v = raw_v.set_index(raw_v.columns[0])
    raw_v = raw_v.loc[list(metrics), list(banks)]

    formats = dict(formats_items)

    numeric_rows = [[parse_value(v) for v in raw_v.loc[m]] for m in raw_v.index]
    num_v = pd.DataFrame(numeric_rows, index=raw_v.index,
                         columns=raw_v.columns).astype("float64")

    return build_pdf_report(raw_v, num_v, formats, source_label,
                            sort_mode, highlight, dates_items, selected_years)


st.set_page_config(
    page_title=REPORT_TITLE,
    page_icon="🏦",
    layout="wide",
)

st.title(f"🏦 {REPORT_TITLE}")

st.caption(f"{SCALE_NOTE} {CONSOLIDATED_NOTE}  \n{APP_VERSION}")

with st.sidebar:
    st.header("Data")

    source_choice = st.radio(
        "Source",
        ["Bundled Excel file", "Upload Excel file"],
        index=0,
    )

    upload = None

    if source_choice == "Upload Excel file":
        upload = st.file_uploader(
            "Excel file: first column = metrics, other columns = banks",
            type=["xlsx"],
        )

    if st.button("🔄 Reload data", use_container_width=True):
        st.rerun()


try:
    if upload is not None:
        raw = load_raw(io.BytesIO(upload.getvalue()))
        source_label = upload.name
    else:
        raw = load_raw(EXCEL_PATH)
        source_label = os.path.basename(EXCEL_PATH)

except FileNotFoundError:
    st.error(
        "Could not find `DTMC stats.xlsx`. Make sure it is in the same GitHub repository folder as `DTMC.py`."
    )
    st.stop()

except Exception as exc:
    st.error(f"Could not read the Excel file: {exc}")
    st.stop()


if raw.empty or raw.shape[1] == 0:
    st.warning("The Excel file has no bank columns.")
    st.stop()


# Pull the per-bank reporting-date row out of the metrics.
raw, report_dates = split_date_row(raw)

# Apply the code-level reporting-date overrides.
if REPORTING_DATE_OVERRIDES:
    if report_dates is None:
        report_dates = pd.Series("", index=raw.columns, dtype=object)
    for b, d in REPORTING_DATE_OVERRIDES.items():
        if b in report_dates.index:
            report_dates[b] = d

numeric, formats = build_numeric(raw)

all_banks = list(raw.columns)
all_metrics = list(raw.index)

with st.sidebar:
    st.header("Filters")

    all_regions = sorted({bank_region(b) for b in all_banks})

    sel_regions = st.multiselect(
        "Regions",
        all_regions,
        default=all_regions,
    )

    region_banks = [b for b in all_banks if bank_region(b) in sel_regions]

    sel_banks = st.multiselect(
        "Banks",
        region_banks,
        default=region_banks,
    )

    sel_metrics = st.multiselect(
        "Metrics",
        all_metrics,
        default=all_metrics,
    )

    selected_years = st.multiselect(
        "Years shown in graphs",
        ["2024", "2025"],
        default=["2024", "2025"],
        help="Select one year to show it alone, or select both years to compare them side by side.",
    )

    st.divider()

    sort_mode = st.radio(
        "Bank order in charts",
        ["By value", "By region, then value", "File order"],
        index=0,
    )

    highlight = st.selectbox(
        "Highlight a bank",
        ["(none)"] + sel_banks,
        index=0,
    )



if not sel_banks or not sel_metrics:
    st.info("Pick at least one bank and one metric.")
    st.stop()

if not selected_years:
    st.info("Select at least one graph year: 2024 or 2025.")
    st.stop()


raw_v = raw.loc[sel_metrics, sel_banks]
num_v = numeric.loc[sel_metrics, sel_banks]

# Reporting dates shown under each tab: one clear line per date group.
_date_lines = date_caption_lines(report_dates, sel_banks)
as_of_md = None
if _date_lines:
    as_of_md = "**Reporting date**  \n" + "  \n".join(
        f"**{line.split(' — ')[0]}** — {line.split(' — ', 1)[1]}"
        for line in _date_lines
    )


def show_reporting_dates():
    if as_of_md:
        with st.container(border=True):
            st.markdown(as_of_md)

with st.sidebar:
    st.divider()
    st.header("Report")
    st.caption(
        "Exports the table and all charts (grouped by topic, one landscape "
        "page each) for the currently selected banks and metrics, using the "
        "current sort order and highlight."
    )

    try:
        report_bytes = get_report_bytes(
            raw_v.reset_index().to_csv(index=False),
            tuple((m, formats[m]) for m in sel_metrics),
            tuple(sel_metrics),
            tuple(sel_banks),
            source_label,
            sort_mode,
            highlight,
            tuple((b, str(report_dates[b])) for b in sel_banks)
            if report_dates is not None else (),
            tuple(selected_years),
            APP_VERSION,
        )

        st.download_button(
            "📄 Download report (PDF)",
            data=report_bytes,
            file_name=f"dtmc_stats_report_{datetime.now():%Y%m%d}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    except ModuleNotFoundError as exc:
        st.warning(
            f"PDF export needs an extra package: `{exc.name}`. "
            "Install with `pip install reportlab matplotlib` "
            "(add both to requirements.txt if deploying)."
        )

    except Exception as exc:
        st.error(f"Could not build the PDF: {exc}")

tab_charts, tab_table, tab_heat = st.tabs([
    "📊 Charts",
    "📋 Table",
    "🌡 Heatmap",
])


with tab_table:
    show_reporting_dates()

    display = pd.DataFrame(
        index=display_labels(raw_v.index),
        columns=raw_v.columns,
        dtype=object,
    )

    for i, m in enumerate(raw_v.index):
        for j, b in enumerate(raw_v.columns):
            display.iloc[i, j] = format_value(
                num_v.iloc[i, j],
                formats[m],
                raw_v.iloc[i, j],
            )

    # Keep the reporting dates visible as the last row of the table.
    if report_dates is not None:
        display.loc["Date"] = [long_date_label(report_dates[b])
                               for b in raw_v.columns]

    neg_mask = num_v.lt(0).values

    def _style(_):
        css = pd.DataFrame("", index=display.index, columns=display.columns)
        # Zebra striping to match the theme.
        for i in range(len(display.index)):
            if i % 2 == 1:
                css.iloc[i, :] = f"background-color: {STRIPE_BG};"
        # Negative values in red.
        for i in range(neg_mask.shape[0]):
            for j in range(neg_mask.shape[1]):
                if neg_mask[i, j]:
                    css.iloc[i, j] += f" color: {NEG}; font-weight: 600;"
        # Date row in muted italics.
        if report_dates is not None:
            css.iloc[-1, :] += " color: #6b7680; font-style: italic;"
        return css

    st.dataframe(
        display.style.apply(_style, axis=None),
        use_container_width=True,
    )

    st.caption(f"{SCALE_NOTE} {CONSOLIDATED_NOTE}")


with tab_heat:
    show_reporting_dates()

    norm = num_v.copy()
    for m, b in CHART_EXCLUSIONS:
        if m in norm.index and b in norm.columns:
            norm.loc[m, b] = float("nan")

    for i, _ in enumerate(norm.index):
        row = norm.iloc[i]
        lo = row.min()
        hi = row.max()

        if pd.isna(lo) or pd.isna(hi) or hi == lo:
            norm.iloc[i] = 0.5
        else:
            norm.iloc[i] = (row - lo) / (hi - lo)

    heat = go.Figure(
        go.Heatmap(
            z=norm.values,
            x=list(norm.columns),
            y=[display_metric_name(m) for m in norm.index],
            colorscale="Teal",
            zmin=0,
            zmax=1,
            xgap=2,
            ygap=2,
            hovertemplate="%{y}<br>%{x}<br>rank score: %{z:.1f}<extra></extra>",
            colorbar=dict(title="relative", thickness=12),
        )
    )

    heat.update_layout(
        height=60 + 42 * len(norm.index),
        margin=dict(l=10, r=10, t=10, b=10),
        yaxis=dict(autorange="reversed"),
        plot_bgcolor="rgba(0,0,0,0)",
    )

    st.plotly_chart(heat, use_container_width=True)


def _metric_bar_fig(chart_group):
    """Create a single-year or grouped FY2024/FY2025 Plotly chart."""
    metric_map = chart_group["metrics"]
    title = chart_group["title"]
    metric_items = list(metric_map.items())

    if not metric_items:
        return None

    first_metric = metric_items[0][1]
    fmt = formats[first_metric]

    # Use all banks that have at least one value in the selected year(s).
    available = []
    for bank in num_v.columns:
        if any(pd.notna(num_v.loc[m, bank]) for _, m in metric_items):
            available.append(bank)

    # Apply metric/bank exclusions. This mainly affects the LBC Net Income YoY row.
    available = [
        bank for bank in available
        if not all(bank in chart_excluded_banks(m) for _, m in metric_items)
    ]

    if not available:
        return None

    # Sort using FY2025 when available, otherwise the latest selected year.
    reference_label, reference_metric = metric_items[-1]
    reference = num_v.loc[reference_metric, available]

    if sort_mode == "By value":
        order = reference.sort_values(ascending=False, na_position="last").index.tolist()
    elif sort_mode == "By region, then value":
        order = sorted(
            available,
            key=lambda b: (
                bank_region(b),
                -(reference.get(b) if pd.notna(reference.get(b)) else float("-inf")),
            ),
        )
    else:
        order = available

    # Title: base metric name, plus a small scale note for currency metrics.
    title_text = display_metric_name(title)
    if fmt == "currency":
        title_text += "<br><sup style='color:#8a949c'>in $ millions (USD)</sup>"

    if fmt == "currency":
        axfmt = "$,.0f"
    elif fmt == "percent":
        axfmt = ".1f"
    else:
        axfmt = ",.1f"

    fig = go.Figure()
    multiple_years = len(metric_items) > 1
    has_year_series = any(bool(year_label) for year_label, _ in metric_items)

    for year_label, metric in metric_items:
        values = num_v.loc[metric, order]
        valid_mask = values.notna()
        trace_banks = [b for b, ok in zip(order, valid_mask) if ok]
        trace_values = [v for v in values if pd.notna(v)]

        if not trace_banks:
            continue

        if multiple_years:
            base_color = YEAR_COLORS.get(year_label, ACCENT)
            colors = [base_color] * len(trace_values)
        elif is_cds_metric(metric):
            colors = [NEG if v > 0 else (POS if v < 0 else ACCENT)
                      for v in trace_values]
        else:
            colors = [NEG if v < 0 else ACCENT for v in trace_values]

        if highlight != "(none)":
            colors = [
                HILITE if bank == highlight else color
                for bank, color in zip(trace_banks, colors)
            ]

        if fmt == "currency":
            hover_value = "$%{y:,.0f}"
        elif fmt == "percent":
            hover_value = "%{y:.1f}%"
        else:
            hover_value = "%{y:,.1f}"

        year_text = f" · FY{year_label}" if year_label else ""
        hovertemplate = (
            "<b>%{customdata[0]}</b> · %{customdata[1]}"
            + year_text
            + "<br>" + hover_value + "<extra></extra>"
        )

        fig.add_trace(go.Bar(
            name=f"FY{year_label}" if year_label else title,
            x=[short_name(b) for b in trace_banks],
            y=trace_values,
            marker_color=colors,
            text=[compact_label(v, fmt) for v in trace_values],
            textposition="outside",
            customdata=[(b, bank_region(b)) for b in trace_banks],
            hovertemplate=hovertemplate,
            cliponaxis=False,
        ))

    if not fig.data:
        return None


    fig.update_layout(
        title=dict(text=title_text, font=dict(size=14)),
        height=320,
        margin=dict(l=10, r=10, t=58, b=10),
        yaxis=dict(tickformat=axfmt, title="", gridcolor=GRID_CLR,
                   zeroline=True, zerolinecolor=ZERO_CLR),
        xaxis=dict(title="", categoryorder="array",
                   categoryarray=[short_name(b) for b in order]),
        plot_bgcolor="rgba(0,0,0,0)",
        bargap=0.22,
        bargroupgap=0.08,
        barmode="group",
        showlegend=has_year_series,
        legend=dict(orientation="h", yanchor="bottom", y=1.01,
                    xanchor="right", x=1),
    )

    return fig


with tab_charts:
    chart_groups = build_chart_groups(sel_metrics, selected_years)

    topic_groups = {
        TOPIC_PERFORMANCE: [],
        TOPIC_RISK: [],
        TOPIC_MARKET: [],
    }

    for group in chart_groups:
        representative = next(iter(group["metrics"].values()))
        name = clean_metric_name(representative).lower()

        if any(k in name for k in MARKET_KEYWORDS):
            topic = TOPIC_MARKET
        elif (formats.get(representative) == "currency"
              or any(k in name for k in PERFORMANCE_KEYWORDS)):
            topic = TOPIC_PERFORMANCE
        else:
            topic = TOPIC_RISK

        topic_groups[topic].append(group)

    topic_pages = [(topic, groups) for topic, groups in topic_groups.items()
                   if groups]

    if not topic_pages:
        st.info("No chartable metrics.")
    else:
        page_tabs = st.tabs([
            f"{TOPIC_ICONS.get(topic, '📊')} {topic} ({len(groups)})"
            for topic, groups in topic_pages
        ])

        for page_tab, (topic, page_groups) in zip(page_tabs, topic_pages):
            with page_tab:
                if topic == TOPIC_MARKET:
                    with st.container(border=True):
                        st.markdown(MARKET_NOTE)
                else:
                    show_reporting_dates()

                selected_label = " and ".join(f"FY{y}" for y in selected_years)
                if len(selected_years) > 1:
                    st.caption(f"{selected_label} are displayed side by side. "
                               "The legend identifies each year by colour.")
                else:
                    st.caption(f"{selected_label} is displayed. "
                               "The legend identifies the selected year by colour.")

                cols = st.columns(2)
                shown = 0

                for group in page_groups:
                    fig = _metric_bar_fig(group)
                    if fig is None:
                        continue

                    cols[shown % 2].plotly_chart(fig, use_container_width=True)
                    shown += 1

                if shown == 0:
                    st.info("No numeric values to chart for this topic.")


st.divider()

st.markdown("##### Bank legend")

legend_cols = st.columns(3)

for i, b in enumerate(sel_banks):
    legend_cols[i % 3].markdown(f"**{b}** — {full_bank_name(b)}")
