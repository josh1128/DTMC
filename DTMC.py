from __future__ import annotations

import os
import io
import math
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

EXCEL_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "DTMC stats.xlsx"
)

METRIC_COL = "In MM (CAD)"

MISSING_TOKENS = {"", "na", "n/a", "n.a.", "-", "—", "nm", "nmf"}

PERFORMANCE_KEYWORDS = (
    "revenue", "income", "earnings", "profit", "margin",
    "yoy", "growth", "eps", "roe", "roa"
)

TOPIC_PERFORMANCE = "Financial Performance"
TOPIC_RISK = "Capital, Liquidity & Credit Quality"

ACCENT = "#1f6f8b"
POS = "#2a9d4a"
NEG = "#c0392b"


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
        if abs(value) <= 1:
            value = value * 100
        return f"{value:.2f}%"

    return f"{value:,.2f}"


def group_metrics_by_topic(metrics, formats):
    perf, risk = [], []

    for m in metrics:
        name = clean_metric_name(m).lower()

        if formats.get(m) == "currency" or any(k in name for k in PERFORMANCE_KEYWORDS):
            perf.append(m)
        else:
            risk.append(m)

    if metrics and (not perf or not risk):
        half = math.ceil(len(metrics) / 2)
        perf, risk = list(metrics[:half]), list(metrics[half:])

    return [(TOPIC_PERFORMANCE, perf), (TOPIC_RISK, risk)]


def load_raw(source):
    df = pd.read_excel(
        source,
        dtype=str,
        keep_default_na=False,
        engine="openpyxl"
    ).fillna("")

    df = df.rename(columns={df.columns[0]: METRIC_COL})
    df = df.set_index(METRIC_COL)
    df.index = make_unique_index(df.index)

    return df


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


st.set_page_config(
    page_title="DTMC Stats Dashboard",
    page_icon="🏦",
    layout="wide",
)

st.title("🏦 DTMC Stats Dashboard")

st.caption(
    "This dashboard reads data from DTMC stats.xlsx. "
    "The first column contains metrics, and the remaining columns contain banks."
)

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
        raw = load_raw(upload)
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
    st.info("Pick at least one bank and one metric.")
    st.stop()


raw_v = raw.loc[sel_metrics, sel_banks]
num_v = numeric.loc[sel_metrics, sel_banks]

tab_charts, tab_table, tab_heat = st.tabs([
    "📊 Charts",
    "📋 Table",
    "🌡 Heatmap",
])


with tab_table:
    display = pd.DataFrame(
        index=[clean_metric_name(m) for m in raw_v.index],
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

    st.dataframe(
        display,
        use_container_width=True,
    )


with tab_heat:
    norm = num_v.copy()

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
            y=[clean_metric_name(m) for m in norm.index],
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


def _metric_bar_fig(metric, fmt):
    row_position = list(num_v.index).index(metric)
    series = num_v.iloc[row_position].dropna()

    if series.empty:
        return None

    if fmt == "percent" and abs(series).max() <= 1:
        series = series * 100

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
        title=dict(text=clean_metric_name(metric), font=dict(size=14)),
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
        st.info("No chartable metrics.")
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

with st.expander("➕ How to update the dashboard"):
    st.markdown(
        """
The dashboard is driven by **DTMC stats.xlsx**.

- Put **DTMC stats.xlsx** in the same GitHub repository folder as `DTMC.py`.
- The first column should contain metrics.
- The remaining columns should contain banks.
- Add a new column to add a bank.
- Add a new row to add a metric.
- Use `$` for currency and `%` for percentage values.
- Use `NA`, `N/A`, `-`, or text like `Meets Req` for unavailable values.
        """
    )
