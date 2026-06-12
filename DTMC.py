
"""
Bank Metrics Dashboard (Streamlit)
==================================
 
A fully data-driven dashboard. It reads a CSV where the FIRST column is the
metric name and every OTHER column is an entity (bank). It then:
 
  * parses values generically ($, %, commas, negatives, "NA", "Meets Req", ...)
  * auto-detects each metric's format (currency / percent / number)
  * renders a formatted table, a normalized heatmap, and one chart per metric
 
Because nothing about the banks or the metrics is hardcoded, ADDING A NEW ROW
(metric) or a NEW COLUMN (bank) to the CSV and reloading the app makes it show
up everywhere automatically.
 
Run with:  streamlit run app.py
"""
 
from __future__ import annotations
 
import os
import io
import math
from datetime import datetime
 
import pandas as pd
import plotly.express as px
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
 
 
# --------------------------------------------------------------------------- #
# PDF report (reportlab for layout, matplotlib for static charts)
# --------------------------------------------------------------------------- #
ACCENT = "#1f6f8b"   # brand teal (also used by the on-screen charts)
POS = "#2a9d4a"
NEG = "#c0392b"
 
 
def _metric_png(metric: str, num_v: pd.DataFrame, fmt: str) -> bytes | None:
    """Render one metric's bar chart to PNG bytes with matplotlib (Agg)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
 
    series = num_v.loc[metric].dropna().sort_values(ascending=False)
    if series.empty:
        return None
    labels, values = list(series.index), series.values
    diverging = (values < 0).any()
    bar_colors = [POS if v >= 0 else NEG for v in values] if diverging else [ACCENT] * len(values)
 
    fig, ax = plt.subplots(figsize=(5.0, 3.0), dpi=150)
    bars = ax.bar(labels, values, color=bar_colors)
    ax.set_title(metric, fontsize=10, fontweight="bold")
    ax.axhline(0, color="#888888", linewidth=0.6)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="x", labelrotation=45, labelsize=7)
    for lbl in ax.get_xticklabels():
        lbl.set_ha("right")
    ax.tick_params(axis="y", labelsize=7)
 
    if fmt == "currency":
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
        labeller = lambda v: f"${v:,.0f}"
    elif fmt == "percent":
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}%"))
        labeller = lambda v: f"{v:g}%"
    else:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.2f}"))
        labeller = lambda v: f"{v:,.2f}"
    ax.bar_label(bars, labels=[labeller(v) for v in values], fontsize=6, padding=2)
 
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()
 
 
def build_pdf_report(raw_v: pd.DataFrame, num_v: pd.DataFrame,
                     formats: dict, source_label: str) -> bytes:
    """Assemble a landscape PDF: title, comparison table, one chart per metric."""
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, Image)
 
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
    header = [Paragraph("Bank", h_cell)] + [Paragraph(m, h_cell) for m in metrics]
    table_data = [header]
    red_cells = []  # (col, row) coords of negative values
    for r, bank in enumerate(banks, start=1):
        cells = [Paragraph(bank, row_lbl)]
        for c, metric in enumerate(metrics, start=1):
            val = num_v.loc[metric, bank]
            cells.append(format_value(val, formats[metric], raw_v.loc[metric, bank]))
            if pd.notna(val) and val < 0:
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
        style.append(("TEXTCOLOR", (c, r), (c, r), colors.HexColor(NEG)))
    table.setStyle(TableStyle(style))
    story += [table, Spacer(1, 16),
              Paragraph("Per-metric comparison", styles["Heading2"]),
              Spacer(1, 4)]
 
    # --- Charts: two per row.
    img_w = (avail_w - 0.2 * inch) / 2
    img_h = img_w * 0.6
    pair = []
    chart_rows = []
    for metric in metrics:
        png = _metric_png(metric, num_v, formats[metric])
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
def get_report_bytes(raw_csv: str, formats_items: tuple,
                     metrics: tuple, banks: tuple, source_label: str) -> bytes:
    """Cached wrapper so the PDF only rebuilds when data/selection changes."""
    raw_v = pd.read_csv(io.StringIO(raw_csv), dtype=str,
                        keep_default_na=False).fillna("").set_index(METRIC_COL)
    raw_v = raw_v.loc[list(metrics), list(banks)]
    formats = dict(formats_items)
    num_v = raw_v.apply(lambda row: [parse_value(v) for v in row], axis=1,
                        result_type="expand").astype("float64")
    num_v.columns, num_v.index = raw_v.columns, raw_v.index
    return build_pdf_report(raw_v, num_v, formats, source_label)
 
 
# --------------------------------------------------------------------------- #
# Data loading
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
 
numeric, formats = build_numeric(raw)
 
all_banks = list(raw.columns)
all_metrics = list(raw.index)
 
# ---- Sidebar: filters ------------------------------------------------------ #
with st.sidebar:
    st.header("Filters")
    sel_banks = st.multiselect("Banks (columns)", all_banks, default=all_banks)
    sel_metrics = st.multiselect("Metrics (rows)", all_metrics, default=all_metrics)
    st.divider()
    sort_charts = st.checkbox("Sort bars by value", value=True)
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
    st.caption("Exports the table and charts for the currently selected "
               "banks and metrics.")
    try:
        report_bytes = get_report_bytes(
            raw_v.reset_index().to_csv(index=False),
            tuple((m, formats[m]) for m in sel_metrics),
            tuple(sel_metrics),
            tuple(sel_banks),
            source_label,
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
 
    neg_mask = num_v.lt(0)
 
    def _style(_):
        css = pd.DataFrame("", index=display.index, columns=display.columns)
        css[neg_mask] = f"color: {NEG}; font-weight: 600;"
        return css
 
    st.dataframe(display.style.apply(_style, axis=None), use_container_width=True)
    st.caption("Negative values are shown in red. Non-numeric entries "
               "(e.g. “Meets Req”) are kept as-is.")
 
# ----- Heatmap -------------------------------------------------------------- #
with tab_heat:
    st.caption("Each metric (row) is min-max normalized across the selected banks "
               "so different scales are comparable. Darker = higher within that row.")
    # Normalize per row 0..1; rows with no spread -> 0.5
    norm = num_v.copy()
    for m in norm.index:
        row = norm.loc[m]
        lo, hi = row.min(), row.max()
        norm.loc[m] = 0.5 if (pd.isna(lo) or pd.isna(hi) or hi == lo) else (row - lo) / (hi - lo)
 
    heat = go.Figure(
        go.Heatmap(
            z=norm.values,
            x=list(norm.columns),
            y=list(norm.index),
            colorscale="Teal",
            zmin=0, zmax=1,
            hovertemplate="%{y}<br>%{x}<br>rank score: %{z:.2f}<extra></extra>",
            colorbar=dict(title="rel."),
        )
    )
    heat.update_layout(height=60 + 42 * len(norm.index),
                       margin=dict(l=10, r=10, t=10, b=10),
                       yaxis=dict(autorange="reversed"))
    st.plotly_chart(heat, use_container_width=True)
 
# ----- Charts (one per metric) --------------------------------------------- #
with tab_charts:
    cols = st.columns(2)
    for i, metric in enumerate(sel_metrics):
        fmt = formats[metric]
        series = num_v.loc[metric].dropna()
        if series.empty:
            continue
 
        frame = series.reset_index()
        frame.columns = ["Bank", "Value"]
        if sort_charts:
            frame = frame.sort_values("Value", ascending=False)
 
        diverging = (frame["Value"] < 0).any()  # YoY-style metrics
        if diverging:
            colors = [POS if v >= 0 else NEG for v in frame["Value"]]
        else:
            colors = [ACCENT] * len(frame)
        if highlight != "(none)":
            colors = ["#e08a1e" if b == highlight else c
                      for b, c in zip(frame["Bank"], colors)]
 
        if fmt == "currency":
            texttmpl, hovertmpl, axfmt = "$%{y:,.0f}", "%{x}<br>$%{y:,.0f}<extra></extra>", "$,.0f"
        elif fmt == "percent":
            texttmpl, hovertmpl, axfmt = "%{y:.2f}%", "%{x}<br>%{y:.2f}%<extra></extra>", ".1f"
        else:
            texttmpl, hovertmpl, axfmt = "%{y:,.2f}", "%{x}<br>%{y:,.2f}<extra></extra>", ",.2f"
 
        fig = go.Figure(go.Bar(
            x=frame["Bank"], y=frame["Value"],
            marker_color=colors, text=frame["Value"],
            texttemplate=texttmpl, textposition="outside",
            hovertemplate=hovertmpl, cliponaxis=False,
        ))
        fig.update_layout(
            title=dict(text=metric, font=dict(size=15)),
            height=360, margin=dict(l=10, r=10, t=46, b=10),
            yaxis=dict(tickformat=axfmt, title=""), xaxis=dict(title=""),
            showlegend=False,
        )
        cols[i % 2].plotly_chart(fig, use_container_width=True)
 
# --------------------------------------------------------------------------- #
st.divider()
with st.expander("➕ How to add a bank or a metric"):
    st.markdown(
        f"""
The dashboard is driven entirely by **`bank_metrics.csv`** (or whatever CSV you
upload), so it stays in sync automatically:
 
* **Add a bank** → add a new **column** (header = bank name), fill in its cells,
  save, then click **🔄 Reload data**. It appears in the table, heatmap and a bar
  in every chart.
* **Add a metric** → add a new **row** under the `{METRIC_COL}` column. Use `$`
  for currency and `%` for percentage values — the app detects the format and a
  new chart is generated for it.
* **Missing / qualitative values** → use `NA`, `N/A`, `-`, or free text like
  `Meets Req`; these are skipped in charts and shown verbatim in the table.
 
No code changes needed for new rows or columns.
        """
    )
 
