"""
Bank Metrics Dashboard (Streamlit)
==================================

A fully data-driven dashboard. It reads a CSV where the FIRST column is the
metric name and every OTHER column is an entity (bank). It then:

  * parses values generically ($, %, commas, negatives, "NA", "Meets Req", ...)
  * auto-detects each metric's format (currency / percent / number)
  * auto-detects each metric's DIRECTION (higher-is-better vs lower-is-better)
  * renders a formatted table (best/worst shaded per row), a direction-aware
    heatmap (darker = BETTER), and one chart per metric with peer-median and
    regulatory-threshold reference lines

Charts are grouped into TWO topic pages -- "Financial Performance" and
"Capital, Liquidity & Credit Quality" -- shown in a landscape 2-column grid so
each page fits on screen without scrolling. The PDF report mirrors the same
grouping, sorting, highlighting, and reference lines.

Because nothing about the banks or the metrics is hardcoded, ADDING A NEW ROW
(metric) or a NEW COLUMN (bank) to the CSV and reloading the app makes it show
up everywhere automatically (topic, direction, and thresholds are inferred
from the metric name).

Run with:  streamlit run app.py
"""

from __future__ import annotations

import os
import io
import re
import math
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bank_metrics.csv")
METRIC_COL = "Metric"  # name of the first column in the CSV

# Seed data (transcribed from the source table, in MM CAD). Used only to create
# the CSV the first time the app runs, if no CSV exists yet. After that, the CSV
# is the single source of truth -- edit it freely.
BANKS = ["TD", "RBC", "BNS", "BMO", "CIBC", "NBC", "LBC", "Desjardins",
         "ATB", "BNP Paribas", "Merrill Lynch", "Citibank NA"]

SEED_ROWS = [
    ("Revenue",
     ["$59,180", "$65,717", "$34,216", "$34,683", "$28,897", "$14,023",
      "$884", "$15,620", "$2,433", "$49,081", "$107,422", "$78,734"]),
    ("Revenue YoY (%)",
     ["-6.4%", "5.5%", "8.14%", "6.20%", "7.86%", "10.12%",
      "-4.16%", "11.07%", "17.82%", "2.39%", "2.02%", "4.00%"]),
    ("Net Income",
     ["$14,910", "$22,138", "$9,548", "$9,731", "$9,818", "$4,612",
      "$27", "$3,321", "$542", "$12,491", "$31,733", "$16,027"]),
    ("Net Income YoY (%)",
     ["-27.4%", "8.7%", "22.58%", "11.73%", "16.48%", "14.81%",
      "-80.06%", "14.71%", "56.19%", "2.18%", "4.01%", "12.03%"]),
    ("CET1 Ratio (%)",
     ["14.5%", "13.7%", "13.3%", "13.0%", "13.4%", "13.7%",
      "11.0%", "23.2%", "11.9%", "12.8%", "11.2%", "12.7%"]),
    ("LCR (%)",
     ["142%", "126%", "124%", "128%", "133%", "189%",
      "Meets Req", "167%", "129%", "134%", "113%", "114%"]),
    ("Gross NPAs/Customer Loans + OREO (%)",
     ["0.56%", "0.86%", "0.99%", "1.07%", "0.64%", "1.23%",
      "1.18%", "0.80%", "NA", "2.97%", "0.98%", "1.09%"]),
    ("New Loan Loss Prov/Avg Customer Loans (%)",
     ["0.47%", "0.43%", "0.61%", "0.45%", "0.41%", "0.45%",
      "0.17%", "0.20%", "0.22%", "0.40%", "0.48%", "1.39%"]),
]

MISSING_TOKENS = {"", "na", "n/a", "n.a.", "-", "—", "nm", "nmf"}

# Short display names for chart x-axes (full name still shown on hover).
# Any bank not listed here falls back to its own (CSV column) name.
SHORT_NAMES = {
    "BNP Paribas": "BNP",
    "Merrill Lynch": "ML",
    "Citibank NA": "Citi",
    "Desjardins": "DESJ",
}

# --------------------------------------------------------------------------- #
# Semantics inferred from the metric NAME (data-driven; extend the keyword
# lists to teach the app about new kinds of metrics -- no other code changes).
# --------------------------------------------------------------------------- #
# Topic grouping (drives the 2 chart pages on screen and in the PDF)
PERFORMANCE_KEYWORDS = ("revenue", "income", "earnings", "profit", "margin",
                        "yoy", "growth", "eps", "roe", "roa")
TOPIC_PERFORMANCE = "Financial Performance"
TOPIC_RISK = "Capital, Liquidity & Credit Quality"

# Direction: metrics matching these keywords are LOWER-is-better. Matched on
# word boundaries so e.g. "NCO" never matches inside "Net Income".
LOWER_IS_BETTER_KEYWORDS = ("npa", "npl", "provision", "loan loss", "charge-off",
                            "chargeoff", "delinquen", "efficiency ratio",
                            "cost/income", "nco")

# Regulatory / reference thresholds, matched by keyword in the metric name:
# (keyword, threshold value, label). Direction of "breach" follows the
# metric's own direction (below-threshold is bad for higher-is-better, etc.).
THRESHOLDS = [
    ("cet1", 11.5, "Reg. min ≈11.5% (D-SIB incl. buffers)"),
    ("lcr", 100.0, "Reg. min 100%"),
]


def higher_is_better(metric: str) -> bool:
    name = metric.lower()
    return not any(re.search(rf"(?<![a-z]){re.escape(k)}", name)
                   for k in LOWER_IS_BETTER_KEYWORDS)


def threshold_for(metric: str):
    """Return (value, label) of a reference threshold for this metric, or None."""
    name = metric.lower()
    for key, value, label in THRESHOLDS:
        if key in name:
            return value, label
    return None


def direction_note(metric: str) -> str:
    return "▲ higher is better" if higher_is_better(metric) else "▼ lower is better"


def short_name(bank: str) -> str:
    return SHORT_NAMES.get(bank, bank)


def group_metrics_by_topic(metrics, formats) -> list[tuple[str, list[str]]]:
    """Split metrics into the two topic pages.

    Returns [(topic_name, [metrics...]), ...] preserving metric order. If the
    keyword rules put everything into one bucket, fall back to an even split
    so both pages are still useful.
    """
    perf, risk = [], []
    for m in metrics:
        name = m.lower()
        if formats.get(m) == "currency" or any(k in name for k in PERFORMANCE_KEYWORDS):
            perf.append(m)
        else:
            risk.append(m)

    if metrics and (not perf or not risk):
        half = math.ceil(len(metrics) / 2)
        perf, risk = list(metrics[:half]), list(metrics[half:])

    return [(TOPIC_PERFORMANCE, perf), (TOPIC_RISK, risk)]


# --------------------------------------------------------------------------- #
# Parsing / formatting helpers (pure functions -- easy to test)
# --------------------------------------------------------------------------- #
def parse_value(raw) -> float | None:
    """Turn a raw cell like '$59,180', '-6.4%' or 'Meets Req' into a float.

    Returns None when the cell is missing or non-numeric (e.g. 'Meets Req').
    Accounting-style negatives '(123)' are supported.
    """
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    s = str(raw).strip()
    if s.lower() in MISSING_TOKENS:
        return None

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1]

    s = s.replace("$", "").replace(",", "").replace("%", "").strip()
    if s.startswith("-"):
        negative, s = True, s[1:]

    try:
        value = float(s)
    except ValueError:
        return None  # text such as 'Meets Req'
    return -value if negative else value


def detect_format(raw_values) -> str:
    """Classify a metric as 'currency', 'percent', or 'number' from its cells."""
    cells = [str(v) for v in raw_values if v is not None]
    if any("$" in c for c in cells):
        return "currency"
    if any("%" in c for c in cells):
        return "percent"
    return "number"


def format_value(value: float | None, fmt: str, original: str = "") -> str:
    """Format a parsed number back into a display string for its metric type."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        # Preserve non-numeric original text (e.g. 'Meets Req', 'NA').
        return str(original) if original not in (None, "") else "—"
    if fmt == "currency":
        return f"${value:,.0f}"
    if fmt == "percent":
        return f"{value:g}%"
    return f"{value:,.2f}"


def compact_label(value: float, fmt: str) -> str:
    """Short label for on-bar text: '$107.4B' instead of '$107,422'.

    Currency cells are in $MM, so >= 1,000 MM is shown in billions.
    """
    if fmt == "currency":
        if abs(value) >= 1000:
            return f"${value / 1000:,.1f}B"
        return f"${value:,.0f}"
    if fmt == "percent":
        return f"{value:,.1f}%"
    return f"{value:,.1f}"


# --------------------------------------------------------------------------- #
# Shared color semantics (screen + PDF):
#   teal   = neutral value
#   red    = genuinely bad (negative on a higher-is-better metric, or a
#            regulatory-threshold breach)
#   orange = the user-highlighted bank
# --------------------------------------------------------------------------- #
ACCENT = "#1f6f8b"     # brand teal
BAD = "#c0392b"        # genuinely bad values
HILITE = "#e08a1e"     # highlighted bank
MEDIAN_CLR = "#5b6770"
THRESH_CLR = "#c0392b"
GRID_CLR = "#eef2f4"
BEST_BG = "#e6f4ea"    # table: best-in-row tint
WORST_BG = "#fdecea"   # table: worst-in-row tint


def is_bad_value(value: float, metric: str) -> bool:
    """A value is 'bad' if it's negative on a higher-is-better metric, or it
    breaches a known regulatory threshold in the wrong direction."""
    if pd.isna(value):
        return False
    hib = higher_is_better(metric)
    if hib and value < 0:
        return True
    thr = threshold_for(metric)
    if thr is not None:
        t, _ = thr
        return value < t if hib else value > t
    return False


def bar_colors(banks, values, metric: str, highlight: str) -> list[str]:
    colors = [BAD if is_bad_value(v, metric) else ACCENT for v in values]
    if highlight and highlight != "(none)":
        colors = [HILITE if b == highlight else c for b, c in zip(banks, colors)]
    return colors


# --------------------------------------------------------------------------- #
# PDF report (reportlab for layout, matplotlib for static charts)
# --------------------------------------------------------------------------- #
def _metric_png(metric: str, num_v: pd.DataFrame, fmt: str,
                sort_by_value: bool, highlight: str) -> bytes | None:
    """Render one metric's bar chart to PNG bytes with matplotlib (Agg),
    mirroring the on-screen chart: same colors, sorting, highlight, direction
    note, median line, and threshold line."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    series = num_v.loc[metric].dropna()
    if series.empty:
        return None
    if sort_by_value:
        series = series.sort_values(ascending=False)
    banks = list(series.index)
    values = series.values
    colors = bar_colors(banks, values, metric, highlight)

    fig, ax = plt.subplots(figsize=(5.0, 3.0), dpi=150)
    bars = ax.bar([short_name(b) for b in banks], values, color=colors)
    ax.set_title(f"{metric}\n{direction_note(metric)}", fontsize=9,
                 fontweight="bold", loc="left")
    ax.axhline(0, color="#888888", linewidth=0.6)

    median = series.median()
    ax.axhline(median, color=MEDIAN_CLR, linewidth=0.8, linestyle="--")
    ax.annotate(f"median {compact_label(median, fmt)}",
                xy=(1.0, median), xycoords=("axes fraction", "data"),
                fontsize=6, color=MEDIAN_CLR, ha="right", va="bottom")

    thr = threshold_for(metric)
    if thr is not None:
        t, label = thr
        ax.axhline(t, color=THRESH_CLR, linewidth=0.8, linestyle=":")
        ax.annotate(label, xy=(0.0, t), xycoords=("axes fraction", "data"),
                    fontsize=6, color=THRESH_CLR, ha="left", va="bottom")

    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color=GRID_CLR, linewidth=0.6)
    ax.set_axisbelow(True)
    ax.tick_params(axis="x", labelrotation=45, labelsize=7)
    for lbl in ax.get_xticklabels():
        lbl.set_ha("right")
    ax.tick_params(axis="y", labelsize=7)

    if fmt == "currency":
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
    elif fmt == "percent":
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}%"))
    else:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.2f}"))
    ax.bar_label(bars, labels=[compact_label(v, fmt) for v in values],
                 fontsize=6, padding=2)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def build_pdf_report(raw_v: pd.DataFrame, num_v: pd.DataFrame, formats: dict,
                     source_label: str, sort_by_value: bool, highlight: str) -> bytes:
    """Assemble a LANDSCAPE PDF: title + comparison table on page 1, then the
    charts grouped by topic -- one landscape page per topic (2 pages of charts)."""
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, Image, PageBreak)

    metrics = list(raw_v.index)
    banks = list(raw_v.columns)
    page_w, page_h = landscape(letter)
    margin = 0.5 * inch
    avail_w = page_w - 2 * margin

    styles = getSampleStyleSheet()
    h_cell = ParagraphStyle("hcell", parent=styles["Normal"], fontSize=6.5,
                            leading=8, textColor=colors.white, fontName="Helvetica-Bold")
    row_lbl = ParagraphStyle("rowlbl", parent=styles["Normal"], fontSize=7,
                             leading=8, fontName="Helvetica-Bold")
    subtitle = ParagraphStyle("sub", parent=styles["Normal"], fontSize=8,
                              textColor=colors.HexColor("#666666"))

    story = [
        Paragraph("Bank Metrics Report", styles["Title"]),
        Paragraph(
            f"Generated {datetime.now():%Y-%m-%d %H:%M} &nbsp;|&nbsp; "
            f"{len(banks)} banks &nbsp;|&nbsp; {len(metrics)} metrics "
            f"&nbsp;|&nbsp; source: {source_label}",
            subtitle,
        ),
        Spacer(1, 10),
    ]

    # --- Comparison table (banks as rows so it grows down the page, not off it).
    header = [Paragraph("Bank", h_cell)] + [
        Paragraph(f"{m}<br/><font size=5>{direction_note(m)}</font>", h_cell)
        for m in metrics
    ]
    table_data = [header]
    red_cells = []  # (col, row) coords of bad values
    for r, bank in enumerate(banks, start=1):
        cells = [Paragraph(bank, row_lbl)]
        for c, metric in enumerate(metrics, start=1):
            val = num_v.loc[metric, bank]
            cells.append(format_value(val, formats[metric], raw_v.loc[metric, bank]))
            if pd.notna(val) and is_bad_value(val, metric):
                red_cells.append((c, r))
        table_data.append(cells)

    first_w = 1.0 * inch
    other_w = (avail_w - first_w) / max(len(metrics), 1)
    table = Table(table_data, colWidths=[first_w] + [other_w] * len(metrics),
                  repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(ACCENT)),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f7f9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for c, r in red_cells:
        style.append(("TEXTCOLOR", (c, r), (c, r), colors.HexColor(BAD)))
    table.setStyle(TableStyle(style))
    story.append(table)

    # --- Charts grouped by topic: each topic starts on its own landscape page.
    img_w = (avail_w - 0.2 * inch) / 2
    img_h = img_w * 0.6

    for topic, topic_metrics in group_metrics_by_topic(metrics, formats):
        if not topic_metrics:
            continue
        story += [PageBreak(),
                  Paragraph(topic, styles["Heading2"]),
                  Spacer(1, 4)]

        pair = []
        chart_rows = []
        for metric in topic_metrics:
            png = _metric_png(metric, num_v, formats[metric], sort_by_value, highlight)
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
        canvas.drawString(margin, 0.3 * inch, "Bank Metrics Report — in MM (CAD)")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(letter),
                            leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=0.6 * inch,
                            title="Bank Metrics Report")
    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


@st.cache_data(show_spinner="Building PDF report…")
def get_report_bytes(raw_csv: str, formats_items: tuple, metrics: tuple,
                     banks: tuple, source_label: str,
                     sort_by_value: bool, highlight: str) -> bytes:
    """Cached wrapper so the PDF only rebuilds when data/selection changes."""
    raw_v = pd.read_csv(io.StringIO(raw_csv), dtype=str,
                        keep_default_na=False).fillna("").set_index(METRIC_COL)
    raw_v = raw_v.loc[list(metrics), list(banks)]
    formats = dict(formats_items)
    num_v = raw_v.apply(lambda row: [parse_value(v) for v in row], axis=1,
                        result_type="expand").astype("float64")
    num_v.columns, num_v.index = raw_v.columns, raw_v.index
    return build_pdf_report(raw_v, num_v, formats, source_label,
                            sort_by_value, highlight)


# --------------------------------------------------------------------------- #
# Data loading & validation
# --------------------------------------------------------------------------- #
def ensure_seed_csv() -> None:
    """Create the CSV from SEED_ROWS the first time, if it doesn't exist."""
    if os.path.exists(CSV_PATH):
        return
    data = {METRIC_COL: [name for name, _ in SEED_ROWS]}
    for i, bank in enumerate(BANKS):
        data[bank] = [vals[i] for _, vals in SEED_ROWS]
    pd.DataFrame(data).to_csv(CSV_PATH, index=False)


def load_raw(source) -> pd.DataFrame:
    """Load the CSV into a frame indexed by metric name (cells kept as strings).

    keep_default_na=False so literal text like 'NA' is preserved verbatim rather
    than being silently converted to a missing value by pandas.
    """
    df = pd.read_csv(source, dtype=str, keep_default_na=False).fillna("")
    if df.columns[0] != METRIC_COL:
        df = df.rename(columns={df.columns[0]: METRIC_COL})
    df = df.set_index(METRIC_COL)
    df.index = df.index.str.strip()
    return df


def validate_raw(raw: pd.DataFrame) -> tuple[pd.DataFrame, list[str], list[tuple]]:
    """Sanity-check the loaded frame.

    Returns (possibly-deduped frame, duplicate metric names dropped,
    list of (metric, bank, cell) text cells that won't chart).
    """
    dupes = []
    if raw.index.duplicated().any():
        dupes = sorted(set(raw.index[raw.index.duplicated()]))
        raw = raw[~raw.index.duplicated(keep="first")]

    text_cells = []
    for m in raw.index:
        for b in raw.columns:
            cell = str(raw.loc[m, b]).strip()
            if cell and cell.lower() not in MISSING_TOKENS \
                    and parse_value(cell) is None:
                text_cells.append((m, b, cell))
    return raw, dupes, text_cells


def build_numeric(raw: pd.DataFrame):
    """Return (numeric_frame, formats_dict) derived from the raw string frame."""
    formats = {m: detect_format(raw.loc[m].tolist()) for m in raw.index}
    numeric = raw.apply(lambda row: [parse_value(v) for v in row], axis=1, result_type="expand")
    numeric.columns = raw.columns
    numeric.index = raw.index
    return numeric.astype("float64"), formats


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="Bank Metrics Dashboard", page_icon="🏦", layout="wide")

st.title("🏦 Bank Metrics Dashboard")
st.caption("In MM (CAD) unless the metric name says otherwise. "
           "This view is generated from the data file — add a row (metric) or a "
           "column (bank) and it appears everywhere automatically.")

ensure_seed_csv()

# ---- Sidebar: data source -------------------------------------------------- #
with st.sidebar:
    st.header("Data")
    source_choice = st.radio("Source", ["Bundled file", "Upload a CSV"], index=0)

    upload = None
    if source_choice == "Upload a CSV":
        upload = st.file_uploader(
            "CSV: first column = metric name, other columns = banks", type=["csv"]
        )

    if st.button("🔄 Reload data", use_container_width=True):
        st.rerun()

# ---- Load ------------------------------------------------------------------ #
try:
    if upload is not None:
        raw = load_raw(io.StringIO(upload.getvalue().decode("utf-8")))
        source_label = upload.name
    else:
        raw = load_raw(CSV_PATH)
        source_label = os.path.basename(CSV_PATH)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read the data: {exc}")
    st.stop()

if raw.empty or raw.shape[1] == 0:
    st.warning("The file has no entity (bank) columns yet.")
    st.stop()

raw, dupe_metrics, text_cells = validate_raw(raw)
if dupe_metrics:
    st.warning(f"Duplicate metric rows in the CSV — kept the first occurrence "
               f"of: {', '.join(dupe_metrics)}")
if text_cells:
    with st.expander(f"ℹ️ {len(text_cells)} non-numeric cell(s) won't appear "
                     "in charts (shown verbatim in the table)"):
        st.dataframe(pd.DataFrame(text_cells, columns=["Metric", "Bank", "Cell"]),
                     use_container_width=True, hide_index=True)

numeric, formats = build_numeric(raw)

all_banks = list(raw.columns)
all_metrics = list(raw.index)

# ---- Sidebar: filters ------------------------------------------------------ #
with st.sidebar:
    st.header("Filters")
    sel_banks = st.multiselect("Banks (columns)", all_banks, default=all_banks)
    sel_metrics = st.multiselect("Metrics (rows)", all_metrics, default=all_metrics)
    st.divider()
    sort_mode = st.radio(
        "Bank order in charts",
        ["By value", "Fixed (CSV order)"],
        index=0,
        help="Fixed order keeps every chart's x-axis identical, so you can "
             "track one bank across all charts.",
    )
    sort_by_value = sort_mode == "By value"
    highlight = st.selectbox("Highlight a bank", ["(none)"] + sel_banks, index=0)

if not sel_banks or not sel_metrics:
    st.info("Pick at least one bank and one metric in the sidebar.")
    st.stop()

raw_v = raw.loc[sel_metrics, sel_banks]
num_v = numeric.loc[sel_metrics, sel_banks]

# ---- Sidebar: PDF report --------------------------------------------------- #
with st.sidebar:
    st.divider()
    st.header("Report")
    st.caption("Exports the table and charts (grouped by topic, one landscape "
               "page each) using the current selection, sort order, and "
               "highlight.")
    try:
        report_bytes = get_report_bytes(
            raw_v.reset_index().to_csv(index=False),
            tuple((m, formats[m]) for m in sel_metrics),
            tuple(sel_metrics),
            tuple(sel_banks),
            source_label,
            sort_by_value,
            highlight,
        )
        st.download_button(
            "📄 Download report (PDF)",
            data=report_bytes,
            file_name=f"bank_metrics_report_{datetime.now():%Y%m%d}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except ModuleNotFoundError as exc:
        st.warning(f"PDF export needs an extra package: `{exc.name}`. "
                   "Install with `pip install reportlab matplotlib`.")
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not build the PDF: {exc}")

# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
tab_charts, tab_table, tab_heat = st.tabs(["📊 Charts", "📋 Table", "🌡 Heatmap"])

# ----- Table ---------------------------------------------------------------- #
with tab_table:
    display = pd.DataFrame(index=raw_v.index, columns=raw_v.columns, dtype=object)
    for m in raw_v.index:
        for b in raw_v.columns:
            display.loc[m, b] = format_value(num_v.loc[m, b], formats[m], raw_v.loc[m, b])
    # Direction note appended to the row label so the table is self-explanatory.
    display.index = [f"{m}  ({direction_note(m)})" for m in raw_v.index]

    bad_mask = pd.DataFrame(
        [[pd.notna(num_v.loc[m, b]) and is_bad_value(num_v.loc[m, b], m)
          for b in raw_v.columns] for m in raw_v.index],
        index=display.index, columns=raw_v.columns,
    )

    def _style(_):
        css = pd.DataFrame("", index=display.index, columns=display.columns)
        css[bad_mask] = f"color: {BAD}; font-weight: 600;"
        # Best / worst per row, respecting the metric's direction.
        for m, row_lbl in zip(raw_v.index, display.index):
            row = num_v.loc[m].dropna()
            if len(row) < 2:
                continue
            best = row.idxmax() if higher_is_better(m) else row.idxmin()
            worst = row.idxmin() if higher_is_better(m) else row.idxmax()
            css.loc[row_lbl, best] += f" background-color: {BEST_BG};"
            css.loc[row_lbl, worst] += f" background-color: {WORST_BG};"
        return css

    st.dataframe(display.style.apply(_style, axis=None), use_container_width=True)
    st.caption(f"🟩 best in row · 🟥 worst in row (direction-aware). Values in "
               f"red breach a regulatory threshold or are negative where higher "
               f"is better. Non-numeric entries (e.g. “Meets Req”) are kept as-is.")

# ----- Heatmap -------------------------------------------------------------- #
with tab_heat:
    st.caption("Each metric (row) is min-max normalized across the selected "
               "banks, with lower-is-better metrics inverted — so **darker "
               "always means better**. Hover shows the actual value.")
    norm = num_v.copy()
    for m in norm.index:
        row = norm.loc[m]
        lo, hi = row.min(), row.max()
        if pd.isna(lo) or pd.isna(hi) or hi == lo:
            norm.loc[m] = 0.5
        else:
            scaled = (row - lo) / (hi - lo)
            norm.loc[m] = scaled if higher_is_better(m) else 1 - scaled

    hover_vals = [[format_value(num_v.loc[m, b], formats[m], raw_v.loc[m, b])
                   for b in num_v.columns] for m in num_v.index]

    heat = go.Figure(
        go.Heatmap(
            z=norm.values,
            x=list(norm.columns),
            y=[f"{m} ({direction_note(m)})" for m in norm.index],
            customdata=hover_vals,
            colorscale="Teal",
            zmin=0, zmax=1,
            hovertemplate="%{y}<br>%{x}: <b>%{customdata}</b>"
                          "<br>score: %{z:.2f} (darker = better)<extra></extra>",
            colorbar=dict(title="better →"),
        )
    )
    heat.update_layout(height=60 + 42 * len(norm.index),
                       margin=dict(l=10, r=10, t=10, b=10),
                       yaxis=dict(autorange="reversed"))
    st.plotly_chart(heat, use_container_width=True)


# ----- Charts (2 topic pages, landscape 2-column grid) ---------------------- #
def _metric_bar_fig(metric: str, fmt: str) -> go.Figure | None:
    """Build the on-screen Plotly bar chart for one metric: direction note in
    the title, peer-median dashed line, regulatory-threshold dotted line,
    short bank labels (full name + rank in hover), compact bar labels."""
    series = num_v.loc[metric].dropna()
    if series.empty:
        return None
    if sort_by_value:
        series = series.sort_values(ascending=False)

    banks = list(series.index)
    values = list(series.values)
    colors = bar_colors(banks, values, metric, highlight)

    # Rank within the metric, respecting direction (1 = best).
    ranked = series.rank(ascending=not higher_is_better(metric), method="min")
    n = len(series)
    custom = [[b, f"{int(ranked[b])} of {n}",
               format_value(series[b], fmt)] for b in banks]

    if fmt == "currency":
        axfmt = "$,.0f"
    elif fmt == "percent":
        axfmt = ".1f"
    else:
        axfmt = ",.2f"

    fig = go.Figure(go.Bar(
        x=[short_name(b) for b in banks], y=values,
        marker_color=colors,
        text=[compact_label(v, fmt) for v in values],
        textposition="outside",
        customdata=custom,
        hovertemplate="<b>%{customdata[0]}</b><br>%{customdata[2]}"
                      "<br>rank: %{customdata[1]}<extra></extra>",
        cliponaxis=False,
    ))

    median = series.median()
    fig.add_hline(y=median, line_dash="dash", line_color=MEDIAN_CLR,
                  line_width=1,
                  annotation_text=f"median {compact_label(median, fmt)}",
                  annotation_position="top right",
                  annotation_font=dict(size=10, color=MEDIAN_CLR))
    thr = threshold_for(metric)
    if thr is not None:
        t, label = thr
        fig.add_hline(y=t, line_dash="dot", line_color=THRESH_CLR, line_width=1,
                      annotation_text=label, annotation_position="bottom left",
                      annotation_font=dict(size=10, color=THRESH_CLR))

    fig.update_layout(
        title=dict(
            text=f"{metric}<br><sup style='color:#8a949c'>{direction_note(metric)}</sup>",
            font=dict(size=14),
        ),
        height=300, margin=dict(l=10, r=10, t=52, b=10),
        yaxis=dict(tickformat=axfmt, title="", gridcolor=GRID_CLR, zeroline=True,
                   zerolinecolor="#c9d2d8"),
        xaxis=dict(title=""),
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
    )
    return fig


with tab_charts:
    topic_pages = [(t, ms) for t, ms in group_metrics_by_topic(sel_metrics, formats) if ms]

    if not topic_pages:
        st.info("No chartable metrics in the current selection.")
    else:
        page_tabs = st.tabs([f"{'📈' if t == TOPIC_PERFORMANCE else '🛡️'} {t} "
                             f"({len(ms)})" for t, ms in topic_pages])
        for page_tab, (topic, page_metrics) in zip(page_tabs, topic_pages):
            with page_tab:
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

# --------------------------------------------------------------------------- #
st.divider()
with st.expander("➕ How to add a bank or a metric"):
    st.markdown(
        f"""
The dashboard is driven entirely by **`bank_metrics.csv`** (or whatever CSV you
upload), so it stays in sync automatically:

* **Add a bank** → add a new **column** (header = bank name), fill in its cells,
  save, then click **🔄 Reload data**. It appears in the table, heatmap and a bar
  in every chart. Add a short display alias to `SHORT_NAMES` if the name is long
  (optional — it falls back to the full name).
* **Add a metric** → add a new **row** under the `{METRIC_COL}` column. Use `$`
  for currency and `%` for percentage values — the app detects the format and a
  new chart is generated for it. The metric's **topic page**, **direction**
  (higher/lower is better), and any **regulatory threshold line** are inferred
  from its name via the keyword lists at the top of the file
  (`PERFORMANCE_KEYWORDS`, `LOWER_IS_BETTER_KEYWORDS`, `THRESHOLDS`).
* **Missing / qualitative values** → use `NA`, `N/A`, `-`, or free text like
  `Meets Req`; these are skipped in charts and shown verbatim in the table.

No code changes needed for new rows or columns.
        """
    )
