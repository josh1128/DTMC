from __future__ import annotations

import os
import io
import math
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bank_metrics.csv")
METRIC_COL = "Metric"

BANKS = ["TD", "RBC", "BNS", "BMO", "CIBC", "NBC", "LBC", "Desjardins",
         "ATB", "BNP Paribas", "Merrill Lynch", "Citibank NA"]

SEED_ROWS = [
    ("Revenue", ["$59,180", "$65,717", "$34,216", "$34,683", "$28,897", "$14,023",
                 "$884", "$15,620", "$2,433", "$49,081", "$107,422", "$78,734"]),
    ("Revenue YoY (%)", ["-6.4%", "5.5%", "8.14%", "6.20%", "7.86%", "10.12%",
                         "-4.16%", "11.07%", "17.82%", "2.39%", "2.02%", "4.00%"]),
    ("Net Income", ["$14,910", "$22,138", "$9,548", "$9,731", "$9,818", "$4,612",
                    "$27", "$3,321", "$542", "$12,491", "$31,733", "$16,027"]),
    ("Net Income YoY (%)", ["-27.4%", "8.7%", "22.58%", "11.73%", "16.48%", "14.81%",
                            "-80.06%", "14.71%", "56.19%", "2.18%", "4.01%", "12.03%"]),
    ("CET1 Ratio (%)", ["14.5%", "13.7%", "13.3%", "13.0%", "13.4%", "13.7%",
                        "11.0%", "23.2%", "11.9%", "12.8%", "11.2%", "12.7%"]),
    ("LCR (%)", ["142%", "126%", "124%", "128%", "133%", "189%",
                 "Meets Req", "167%", "129%", "134%", "113%", "114%"]),
    ("Gross NPAs/Customer Loans + OREO (%)", ["0.56%", "0.86%", "0.99%", "1.07%",
                                               "0.64%", "1.23%", "1.18%", "0.80%",
                                               "NA", "2.97%", "0.98%", "1.09%"]),
    ("New Loan Loss Prov/Avg Customer Loans (%)", ["0.47%", "0.43%", "0.61%", "0.45%",
                                                   "0.41%", "0.45%", "0.17%", "0.20%",
                                                   "0.22%", "0.40%", "0.48%", "1.39%"]),
]

MISSING_TOKENS = {"", "na", "n/a", "n.a.", "-", "—", "nm", "nmf"}

PERFORMANCE_KEYWORDS = ("revenue", "income", "earnings", "profit", "margin",
                        "yoy", "growth", "eps", "roe", "roa")

TOPIC_PERFORMANCE = "Financial Performance"
TOPIC_RISK = "Capital, Liquidity & Credit Quality"

ACCENT = "#1f6f8b"
POS = "#2a9d4a"
NEG = "#c0392b"


def group_metrics_by_topic(metrics, formats) -> list[tuple[str, list[str]]]:
    perf, risk = [], []

    for m in metrics:
        name = str(m).lower()
        if formats.get(m) == "currency" or any(k in name for k in PERFORMANCE_KEYWORDS):
            perf.append(m)
        else:
            risk.append(m)

    if metrics and (not perf or not risk):
        half = math.ceil(len(metrics) / 2)
        perf, risk = list(metrics[:half]), list(metrics[half:])

    return [(TOPIC_PERFORMANCE, perf), (TOPIC_RISK, risk)]


def parse_value(raw) -> float | None:
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None

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


def detect_format(raw_values) -> str:
    cells = [str(v) for v in raw_values if v is not None]

    if any("$" in c for c in cells):
        return "currency"

    if any("%" in c for c in cells):
        return "percent"

    return "number"


def format_value(value: float | None, fmt: str, original: str = "") -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return str(original) if original not in (None, "") else "—"

    if fmt == "currency":
        return f"${value:,.0f}"

    if fmt == "percent":
        return f"{value:g}%"

    return f"{value:,.2f}"


def _metric_png(metric: str, num_v: pd.DataFrame, fmt: str) -> bytes | None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    series = num_v.loc[metric].dropna().sort_values(ascending=False)

    if series.empty:
        return None

    labels = list(series.index)
    values = series.values

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
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        Image,
        PageBreak,
    )

    metrics = list(raw_v.index)
    banks = list(raw_v.columns)

    page_w, page_h = landscape(letter)
    margin = 0.5 * inch
    avail_w = page_w - 2 * margin

    styles = getSampleStyleSheet()

    h_cell = ParagraphStyle(
        "hcell",
        parent=styles["Normal"],
        fontSize=6.5,
        leading=8,
        textColor=colors.white,
        fontName="Helvetica-Bold",
    )

    row_lbl = ParagraphStyle(
        "rowlbl",
        parent=styles["Normal"],
        fontSize=7,
        leading=8,
        fontName="Helvetica-Bold",
    )

    subtitle = ParagraphStyle(
        "sub",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.HexColor("#666666"),
    )

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

    header = [Paragraph("Bank", h_cell)] + [Paragraph(m, h_cell) for m in metrics]
    table_data = [header]

    red_cells = []

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

    table = Table(
        table_data,
        colWidths=[first_w] + [other_w] * len(metrics),
        repeatRows=1,
    )

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
    story.append(table)

    img_w = (avail_w - 0.2 * inch) / 2
    img_h = img_w * 0.6

    for topic, topic_metrics in group_metrics_by_topic(metrics, formats):
        if not topic_metrics:
            continue

        story += [
            PageBreak(),
            Paragraph(topic, styles["Heading2"]),
            Spacer(1, 4),
        ]

        pair = []
        chart_rows = []

        for metric in topic_metrics:
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

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=0.6 * inch,
        title="Bank Metrics Report",
    )

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)

    return buf.getvalue()


@st.cache_data(show_spinner="Building PDF report…")
def get_report_bytes(raw_csv: str, formats_items: tuple,
                     metrics: tuple, banks: tuple, source_label: str) -> bytes:
    df = pd.read_csv(
        io.StringIO(raw_csv),
        dtype=str,
        keep_default_na=False,
    ).fillna("")

    # FIX: guarantee the first column is called Metric
    if METRIC_COL not in df.columns:
        df = df.rename(columns={df.columns[0]: METRIC_COL})

    raw_v = df.set_index(METRIC_COL)

    raw_v.index = raw_v.index.astype(str).str.strip()

    raw_v = raw_v.loc[list(metrics), list(banks)]

    formats = dict(formats_items)

    num_v = raw_v.apply(
        lambda row: [parse_value(v) for v in row],
        axis=1,
        result_type="expand",
    ).astype("float64")

    num_v.columns = raw_v.columns
    num_v.index = raw_v.index

    return build_pdf_report(raw_v, num_v, formats, source_label)


def ensure_seed_csv() -> None:
    if os.path.exists(CSV_PATH):
        return

    data = {METRIC_COL: [name for name, _ in SEED_ROWS]}

    for i, bank in enumerate(BANKS):
        data[bank] = [vals[i] for _, vals in SEED_ROWS]

    pd.DataFrame(data).to_csv(CSV_PATH, index=False)


def load_raw(source) -> pd.DataFrame:
    df = pd.read_csv(
        source,
        dtype=str,
        keep_default_na=False,
    ).fillna("")

    # Guarantees first column is Metric, even if CSV uses index/Unnamed/etc.
    if METRIC_COL not in df.columns:
        df = df.rename(columns={df.columns[0]: METRIC_COL})

    df = df.set_index(METRIC_COL)
    df.index = df.index.astype(str).str.strip()

    return df


def build_numeric(raw: pd.DataFrame):
    formats = {m: detect_format(raw.loc[m].tolist()) for m in raw.index}

    numeric = raw.apply(
        lambda row: [parse_value(v) for v in row],
        axis=1,
        result_type="expand",
    )

    numeric.columns = raw.columns
    numeric.index = raw.index

    return numeric.astype("float64"), formats


st.set_page_config(
    page_title="Bank Metrics Dashboard",
    page_icon="🏦",
    layout="wide",
)

st.title("🏦 Bank Metrics Dashboard")

st.caption(
    "In MM (CAD) unless the metric name says otherwise. "
    "This view is generated from the data file — add a row or column and it appears automatically."
)

ensure_seed_csv()

with st.sidebar:
    st.header("Data")

    source_choice = st.radio(
        "Source",
        ["Bundled file", "Upload a CSV"],
        index=0,
    )

    upload = None

    if source_choice == "Upload a CSV":
        upload = st.file_uploader(
            "CSV: first column = metric name, other columns = banks",
            type=["csv"],
        )

    if st.button("🔄 Reload data", use_container_width=True):
        st.rerun()

try:
    if upload is not None:
        raw = load_raw(io.StringIO(upload.getvalue().decode("utf-8")))
        source_label = upload.name
    else:
        raw = load_raw(CSV_PATH)
        source_label = os.path.basename(CSV_PATH)

except Exception as exc:
    st.error(f"Could not read the data: {exc}")
    st.stop()

if raw.empty or raw.shape[1] == 0:
    st.warning("The file has no bank columns yet.")
    st.stop()

numeric, formats = build_numeric(raw)

all_banks = list(raw.columns)
all_metrics = list(raw.index)

with st.sidebar:
    st.header("Filters")

    sel_banks = st.multiselect(
        "Banks",
        all_banks,
        default=all_banks,
    )

    sel_metrics = st.multiselect(
        "Metrics",
        all_metrics,
        default=all_metrics,
    )

    st.divider()

    sort_charts = st.checkbox("Sort bars by value", value=True)

    highlight = st.selectbox(
        "Highlight a bank",
        ["(none)"] + sel_banks,
        index=0,
    )

if not sel_banks or not sel_metrics:
    st.info("Pick at least one bank and one metric in the sidebar.")
    st.stop()

raw_v = raw.loc[sel_metrics, sel_banks]
num_v = numeric.loc[sel_metrics, sel_banks]

with st.sidebar:
    st.divider()
    st.header("Report")

    st.caption(
        "Exports the table and charts for the currently selected banks and metrics."
    )

    try:
        # FIX: ensure reset index column is always called Metric
        csv_df = raw_v.reset_index()
        csv_df = csv_df.rename(columns={csv_df.columns[0]: METRIC_COL})

        report_bytes = get_report_bytes(
            csv_df.to_csv(index=False),
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
        st.warning(
            f"PDF export needs an extra package: `{exc.name}`. "
            "Install with `pip install reportlab matplotlib`."
        )

    except Exception as exc:
        st.error(f"Could not build the PDF: {exc}")


tab_charts, tab_table, tab_heat = st.tabs([
    "📊 Charts",
    "📋 Table",
    "🌡 Heatmap",
])


with tab_table:
    display = pd.DataFrame(
        index=raw_v.index,
        columns=raw_v.columns,
        dtype=object,
    )

    for m in raw_v.index:
        for b in raw_v.columns:
            display.loc[m, b] = format_value(
                num_v.loc[m, b],
                formats[m],
                raw_v.loc[m, b],
            )

    neg_mask = num_v.lt(0)

    def _style(_):
        css = pd.DataFrame("", index=display.index, columns=display.columns)
        css[neg_mask] = f"color: {NEG}; font-weight: 600;"
        return css

    st.dataframe(
        display.style.apply(_style, axis=None),
        use_container_width=True,
    )

    st.caption(
        "Negative values are shown in red. Non-numeric entries like 'Meets Req' are kept as-is."
    )


with tab_heat:
    st.caption(
        "Each metric is min-max normalized across selected banks. Darker = higher within that row."
    )

    norm = num_v.copy()

    for m in norm.index:
        row = norm.loc[m]
        lo = row.min()
        hi = row.max()

        if pd.isna(lo) or pd.isna(hi) or hi == lo:
            norm.loc[m] = 0.5
        else:
            norm.loc[m] = (row - lo) / (hi - lo)

    heat = go.Figure(
        go.Heatmap(
            z=norm.values,
            x=list(norm.columns),
            y=list(norm.index),
            colorscale="Teal",
            zmin=0,
            zmax=1,
            hovertemplate="%{y}<br>%{x}<br>rank score: %{z:.2f}<extra></extra>",
            colorbar=dict(title="rel."),
        )
    )

    heat.update_layout(
        height=60 + 42 * len(norm.index),
        margin=dict(l=10, r=10, t=10, b=10),
        yaxis=dict(autorange="reversed"),
    )

    st.plotly_chart(heat, use_container_width=True)


def _metric_bar_fig(metric: str, fmt: str) -> go.Figure | None:
    series = num_v.loc[metric].dropna()

    if series.empty:
        return None

    frame = series.reset_index()
    frame.columns = ["Bank", "Value"]

    if sort_charts:
        frame = frame.sort_values("Value", ascending=False)

    diverging = (frame["Value"] < 0).any()

    if diverging:
        colors = [POS if v >= 0 else NEG for v in frame["Value"]]
    else:
        colors = [ACCENT] * len(frame)

    if highlight != "(none)":
        colors = [
            "#e08a1e" if b == highlight else c
            for b, c in zip(frame["Bank"], colors)
        ]

    if fmt == "currency":
        texttmpl = "$%{y:,.0f}"
        hovertmpl = "%{x}<br>$%{y:,.0f}<extra></extra>"
        axfmt = "$,.0f"
    elif fmt == "percent":
        texttmpl = "%{y:.2f}%"
        hovertmpl = "%{x}<br>%{y:.2f}%<extra></extra>"
        axfmt = ".1f"
    else:
        texttmpl = "%{y:,.2f}"
        hovertmpl = "%{x}<br>%{y:,.2f}<extra></extra>"
        axfmt = ",.2f"

    fig = go.Figure(
        go.Bar(
            x=frame["Bank"],
            y=frame["Value"],
            marker_color=colors,
            text=frame["Value"],
            texttemplate=texttmpl,
            textposition="outside",
            hovertemplate=hovertmpl,
            cliponaxis=False,
        )
    )

    fig.update_layout(
        title=dict(text=metric, font=dict(size=14)),
        height=300,
        margin=dict(l=10, r=10, t=42, b=10),
        yaxis=dict(tickformat=axfmt, title=""),
        xaxis=dict(title=""),
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
        st.info("No chartable metrics in the current selection.")
    else:
        page_tabs = st.tabs([
            f"{'📈' if t == TOPIC_PERFORMANCE else '🛡️'} {t} ({len(ms)})"
            for t, ms in topic_pages
        ])

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


st.divider()

with st.expander("➕ How to add a bank or a metric"):
    st.markdown(
        f"""
The dashboard is driven entirely by **`bank_metrics.csv`** or the CSV you upload.

- **Add a bank**: add a new column with the bank name as the header.
- **Add a metric**: add a new row under the `{METRIC_COL}` column.
- Use `$` for currency and `%` for percentage values.
- Missing or qualitative values can be `NA`, `N/A`, `-`, or text like `Meets Req`.
- Click **🔄 Reload data** after editing the file.

No code changes are needed for new rows or columns.
        """
    )
