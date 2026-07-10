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
CDS_NOTE = (
    "CDS spreads: a positive change (widening, red) signals deteriorating "
    "credit perception; a negative change (tightening, green) signals "
    "improvement — hence the inverted colors on CDS charts."
)

MARKET_NOTE = (
    "Market indicators are as of June 2026 — the ▲ 3 months change is "
    "measured vs. March 2026, and the ▲ 1 year change vs. June 2025."
)

LBC_NI_YOY_NOTE = (
    "Note: LBC YoY change is not included because its previous year's value "
    "was negative, so it is not economically meaningful."
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
def _metric_png(metric, num_v, fmt, sort_mode, highlight):
    """Render one metric's bar chart to PNG bytes with matplotlib (Agg),
    mirroring the on-screen chart: same colors, ordering, and highlight."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    series = num_v.loc[metric].dropna()
    series = series.drop(index=[b for b in chart_excluded_banks(metric)
                                if b in series.index])

    if series.empty:
        return None

    if sort_mode == "By value":
        series = series.sort_values(ascending=False)
    elif sort_mode == "By region, then value":
        order = sorted(series.index,
                       key=lambda b: (bank_region(b), -series[b]))
        series = series.loc[order]

    banks = list(series.index)
    values = series.values

    if is_cds_metric(metric):
        # Inverted: widening (positive) = red, tightening (negative) = green.
        bar_colors = [NEG if v > 0 else (POS if v < 0 else ACCENT)
                      for v in values]
    else:
        bar_colors = [NEG if v < 0 else ACCENT for v in values]

    if highlight != "(none)":
        bar_colors = [
            HILITE if b == highlight else c
            for b, c in zip(banks, bar_colors)
        ]

    fig, ax = plt.subplots(figsize=(6.4, 3.4), dpi=150)
    bars = ax.bar([short_name(b) for b in banks], values, color=bar_colors)
    ax.grid(axis="y", color=GRID_CLR, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.margins(y=0.18)  # headroom for the staggered labels

    ax.set_title(display_metric_name(metric), fontsize=10, fontweight="bold")
    ax.axhline(0, color="#888888", linewidth=0.6)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="x", labelrotation=45, labelsize=7)
    for lbl in ax.get_xticklabels():
        lbl.set_ha("right")
    ax.tick_params(axis="y", labelsize=7)

    if fmt == "currency":
        ax.set_ylabel("US$ MM", fontsize=7, color="#666666")
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
        fmt_key = "currency"
    elif fmt == "percent":
        ax.yaxis.set_major_formatter(
            FuncFormatter(lambda v, _: f"{v:,.1f}%"))
        fmt_key = "percent"
    else:
        ax.yaxis.set_major_formatter(
            FuncFormatter(lambda v, _: f"{v:,.1f}"))
        fmt_key = "number"

    # Stagger labels on two levels so close-valued neighbours never overlap.
    labels = [compact_label(v, fmt_key) for v in values]
    for parity, pad in ((0, 2), (1, 12)):
        level = [lbl if i % 2 == parity else ""
                 for i, lbl in enumerate(labels)]
        ax.bar_label(bars, labels=level, fontsize=6, padding=pad)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def build_pdf_report(raw_v, num_v, formats, source_label,
                     sort_mode, highlight, dates_items):
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
        Paragraph("DTMC Stats Report", styles["Title"]),
        Paragraph(
            f"Generated {datetime.now():%Y-%m-%d %H:%M} &nbsp;|&nbsp; "
            f"{len(banks)} banks &nbsp;|&nbsp; {len(metrics)} metrics "
            f"&nbsp;|&nbsp; source: {source_label} &nbsp;|&nbsp; {SCALE_NOTE}",
            subtitle,
        ),
        Paragraph(CONSOLIDATED_NOTE, subtitle),
        Spacer(1, 10),
    ]

    # --- Comparison table (banks as rows so it grows down the page).
    header = [Paragraph("Bank", h_cell), Paragraph("Region", h_cell)]
    if dates:
        header.append(Paragraph("As of", h_cell))
    header += [Paragraph(display_metric_name(m), h_cell) for m in metrics]

    n_lead = 3 if dates else 2  # label columns before the metric columns
    table_data = [header]
    red_cells = []  # (col, row) coords of negative values

    for r, bank in enumerate(banks, start=1):
        cells = [Paragraph(bank, row_lbl), bank_region(bank)]
        if dates:
            cells.append(long_date_label(dates.get(bank, "")))
        for c, metric in enumerate(metrics):
            val = num_v.loc[metric, bank]
            cells.append(format_value(val, formats[metric], raw_v.loc[metric, bank]))
            if pd.notna(val) and val < 0:
                red_cells.append((c + n_lead, r))
        table_data.append(cells)

    first_w = 1.0 * inch
    region_w = 0.7 * inch
    date_w = 0.85 * inch if dates else 0.0
    other_w = (avail_w - first_w - region_w - date_w) / max(len(metrics), 1)
    col_widths = ([first_w, region_w] + ([date_w] if dates else [])
                  + [other_w] * len(metrics))
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(ACCENT)),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (n_lead - 1, 1), (-1, -1), "CENTER"),
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
    story += [Spacer(1, 6), Paragraph(legend, subtitle)]

    # --- Charts grouped by topic: each topic starts on its own page.
    img_w = (avail_w - 0.2 * inch) / 2
    img_h = img_w * 0.6

    for topic, topic_metrics in group_metrics_by_topic(metrics, formats):
        if not topic_metrics:
            continue

        story += [PageBreak(),
                  Paragraph(topic, styles["Heading2"])]

        if topic == TOPIC_MARKET:
            story.append(Paragraph(MARKET_NOTE, subtitle))
            story.append(Paragraph(CDS_NOTE, subtitle))
        else:
            for line in date_caption_lines(
                    pd.Series(dates) if dates else None, banks):
                story.append(Paragraph(line, subtitle))
            if topic == TOPIC_PERFORMANCE and "LBC" in banks and any(
                    m in metrics for (m, _) in CHART_EXCLUSIONS):
                story.append(Paragraph(LBC_NI_YOY_NOTE, subtitle))
        story.append(Spacer(1, 4))

        pair = []
        chart_rows = []
        for metric in topic_metrics:
            png = _metric_png(metric, num_v, formats[metric],
                              sort_mode, highlight)
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
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]))
            story.append(grid)

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#999999"))
        canvas.drawRightString(page_w - margin, 0.3 * inch, f"Page {doc.page}")
        canvas.drawString(margin, 0.3 * inch,
                          "DTMC Stats Report — Revenue & Net Income in MM (USD)")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(letter),
                            leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=0.6 * inch,
                            title="DTMC Stats Report")
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


@st.cache_data(show_spinner="Building PDF report…")
def get_report_bytes(raw_csv, formats_items, metrics, banks, source_label,
                     sort_mode, highlight, dates_items):
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
                            sort_mode, highlight, dates_items)


st.set_page_config(
    page_title="DTMC Stats Dashboard",
    page_icon="🏦",
    layout="wide",
)

st.title("🏦 DTMC Stats Dashboard")

st.caption(f"{SCALE_NOTE} {CONSOLIDATED_NOTE}")

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


def _metric_bar_fig(metric, fmt):
    row_position = list(num_v.index).index(metric)
    series = num_v.iloc[row_position].dropna()
    series = series.drop(index=[b for b in chart_excluded_banks(metric)
                                if b in series.index])

    if series.empty:
        return None

    frame = series.reset_index()
    frame.columns = ["Bank", "Value"]
    frame["Region"] = [bank_region(b) for b in frame["Bank"]]

    if sort_mode == "By value":
        frame = frame.sort_values("Value", ascending=False)
    elif sort_mode == "By region, then value":
        frame = frame.sort_values(["Region", "Value"],
                                  ascending=[True, False])
    # "File order": keep as-is

    if is_cds_metric(metric):
        # Inverted: widening (positive) = red, tightening (negative) = green.
        colors = [NEG if v > 0 else (POS if v < 0 else ACCENT)
                  for v in frame["Value"]]
    else:
        colors = [NEG if v < 0 else ACCENT for v in frame["Value"]]

    if highlight != "(none)":
        colors = [
            HILITE if b == highlight else c
            for b, c in zip(frame["Bank"], colors)
        ]

    # Title: metric name, plus a small scale note for currency metrics.
    subline_parts = []
    if fmt == "currency":
        subline_parts.append("in $ millions (USD)")

    title_text = display_metric_name(metric)
    if subline_parts:
        title_text += ("<br><sup style='color:#8a949c'>"
                       + "  ·  ".join(subline_parts) + "</sup>")

    if fmt == "currency":
        hovertmpl = ("<b>%{customdata[0]}</b> · %{customdata[1]}"
                     "<br>$%{y:,.0f}<extra></extra>")
        axfmt = "$,.0f"
    elif fmt == "percent":
        hovertmpl = ("<b>%{customdata[0]}</b> · %{customdata[1]}"
                     "<br>%{y:.1f}%<extra></extra>")
        axfmt = ".1f"
    else:
        hovertmpl = ("<b>%{customdata[0]}</b> · %{customdata[1]}"
                     "<br>%{y:,.1f}<extra></extra>")
        axfmt = ",.1f"

    fig = go.Figure(
        go.Bar(
            x=[short_name(b) for b in frame["Bank"]],
            y=frame["Value"],
            marker_color=colors,
            text=[compact_label(v, fmt) for v in frame["Value"]],
            textposition="outside",
            customdata=list(zip(frame["Bank"], frame["Region"])),
            hovertemplate=hovertmpl,
            cliponaxis=False,
        )
    )

    fig.update_layout(
        title=dict(text=title_text, font=dict(size=14)),
        height=300,
        margin=dict(l=10, r=10, t=48, b=10),
        yaxis=dict(tickformat=axfmt, title="", gridcolor=GRID_CLR,
                   zeroline=True, zerolinecolor=ZERO_CLR),
        xaxis=dict(title=""),
        plot_bgcolor="rgba(0,0,0,0)",
        bargap=0.25,
        showlegend=False,
    )

    return fig


with tab_charts:
    topic_pages = [
        (t, ms)
        for t, ms in group_metrics_by_topic(sel_metrics, formats)
        if ms
    ]

    if not topic_pages:
        st.info("No chartable metrics.")
    else:
        page_tabs = st.tabs([
            f"{TOPIC_ICONS.get(t, '📊')} {t} ({len(ms)})"
            for t, ms in topic_pages
        ])

        for page_tab, (topic, page_metrics) in zip(page_tabs, topic_pages):
            with page_tab:
                if topic == TOPIC_MARKET:
                    with st.container(border=True):
                        st.markdown(f"{MARKET_NOTE}  \n{CDS_NOTE}")
                else:
                    show_reporting_dates()
                    if (topic == TOPIC_PERFORMANCE
                            and "LBC" in sel_banks
                            and any(m in page_metrics
                                    for (m, _) in CHART_EXCLUSIONS)):
                        st.caption(LBC_NI_YOY_NOTE)

                cols = st.columns(2)
                shown = 0

                for metric in page_metrics:
                    fig = _metric_bar_fig(metric, formats[metric])

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
