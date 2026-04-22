# ============================
# One-cell Streamlit launcher (dashboard)
# This cell will:
# 1) Write the Streamlit dashboard code into app_8505.py
# 2) Run the Streamlit app on port 8505
# 3) Open the dashboard using: http://localhost:8505
# ============================

import textwrap
import os
import sys
import subprocess

APP_FILE = "app_8505.py"
PORT = 8505

app_code = r'''
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------
# Streamlit page configuration
# -----------------------------
st.set_page_config(page_title="Cut to Ship Dashboard", layout="wide")

# -----------------------------
# Excel data source
# -----------------------------
EXCEL_PATH = "Cut to Ship Modified Cleaned.xlsx"
SHEET_NAME = "All"

# -----------------------------
# Corporate theme / styling
# -----------------------------
CORPORATE_COLORS = [
    "#0F4C81",
    "#4DA3D9",
    "#17A2B8",
    "#D4A017",
    "#FF6B6B",
    "#7A5AF8",
    "#2FBF71",
    "#F39C12",
    "#6C757D",
    "#E83E8C",
    "#20C997",
    "#6610F2",
]

st.markdown("""
<style>
    .stApp {
        background: linear-gradient(180deg, #f4f7fb 0%, #eef3f8 100%);
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f2747 0%, #163a63 100%);
        color: white;
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    .sidebar-card {
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 16px;
        padding: 16px 16px 8px 16px;
        margin-bottom: 1rem;
    }

    .hero-box {
        background: linear-gradient(135deg, #0f2747 0%, #163a63 60%, #1d4c7f 100%);
        padding: 24px 28px;
        border-radius: 18px;
        color: white;
        box-shadow: 0 12px 28px rgba(15, 39, 71, 0.18);
        margin-bottom: 1rem;
    }

    .hero-title {
        font-size: 2rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }

    .hero-sub {
        font-size: 0.95rem;
        opacity: 0.92;
    }

    .section-box {
        background: #ffffff;
        border-radius: 18px;
        padding: 18px 20px;
        box-shadow: 0 10px 22px rgba(31, 45, 61, 0.08);
        border: 1px solid #dbe5f0;
        margin-bottom: 1rem;
    }

    .metric-card {
        background: #ffffff;
        border-radius: 18px;
        padding: 16px 18px;
        box-shadow: 0 10px 22px rgba(31, 45, 61, 0.08);
        border: 1px solid #dbe5f0;
        min-height: 120px;
        position: relative;
        overflow: hidden;
        margin-bottom: 12px;
    }

    .metric-card::before {
        content: "";
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        height: 5px;
        background: linear-gradient(90deg, #0F4C81, #4DA3D9);
    }

    .metric-label {
        font-size: 0.9rem;
        font-weight: 600;
        color: #486581;
        margin-bottom: 0.45rem;
    }

    .metric-value {
        font-size: 1.95rem;
        font-weight: 800;
        color: #102a43;
        line-height: 1.1;
    }

    .metric-delta {
        margin-top: 0.45rem;
        font-size: 0.9rem;
        font-weight: 700;
    }

    .page-title {
        font-size: 1.9rem;
        font-weight: 800;
        color: #102a43;
        margin-bottom: 0.35rem;
    }

    .page-subtitle {
        color: #486581;
        font-size: 0.95rem;
        margin-bottom: 1rem;
    }

    div[data-testid="stSelectbox"] > div,
    div[data-testid="stMultiSelect"] > div {
        border-radius: 10px !important;
    }

    .small-note {
        color: #6b7c93;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Helper functions
# -----------------------------
MONTH_ORDER = {
    "JAN": 1, "JANUARY": 1,
    "FEB": 2, "FEBRUARY": 2,
    "MAR": 3, "MARCH": 3,
    "APR": 4, "APRIL": 4,
    "MAY": 5,
    "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7,
    "AUG": 8, "AUGUST": 8,
    "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
    "OCT": 10, "OCTOBER": 10,
    "NOV": 11, "NOVEMBER": 11,
    "DEC": 12, "DECEMBER": 12,
}

MONTH_LABEL = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}

def safe_div(n, d):
    if d is None or pd.isna(d) or d == 0:
        return np.nan
    return n / d

def percent_fmt(x, decimals=1):
    return "NA" if pd.isna(x) else f"{x*100:,.{decimals}f}%"

def num_fmt(x):
    return "NA" if pd.isna(x) else f"{x:,.0f}"

def delta_fmt(curr, prev, decimals=1):
    if pd.isna(curr) or pd.isna(prev) or prev == 0:
        return "Base year"
    delta = ((curr - prev) / abs(prev)) * 100
    if delta > 0:
        return f"▲ {abs(delta):,.{decimals}f}% vs previous year"
    elif delta < 0:
        return f"▼ {abs(delta):,.{decimals}f}% vs previous year"
    return f"■ 0.0% vs previous year"

def clean_week(series):
    w = series.astype(str).str.strip()
    w = w.str.replace("Week", "", regex=False).str.replace("W", "", regex=False).str.strip()
    return pd.to_numeric(w, errors="coerce")

def clean_month(series):
    m = series.astype(str).str.upper().str.strip()
    return m.map(MONTH_ORDER)

def pick_display_years(df_source, selected_years):
    if selected_years:
        return sorted(selected_years)
    years = sorted([int(x) for x in df_source["Year"].dropna().unique().tolist()])
    if not years:
        return []
    return [max(years)]

def weekly_totals(df):
    agg = {
        "OrderQty": ("OrderQty", "sum"),
        "CutQty": ("CutQty", "sum"),
        "ShipQty": ("ShipQty", "sum"),
        "CutShipDiff": ("CutShipDiff", "sum"),
    }
    wk = df.groupby(["Year", "Week_Num"], dropna=False, as_index=False).agg(**agg)
    wk["Cut/Ship"] = wk["ShipQty"] / wk["CutQty"]
    wk["Order/Ship"] = wk["ShipQty"] / wk["OrderQty"]
    wk["Order/Cut"] = wk["CutQty"] / wk["OrderQty"]
    wk = wk.replace([np.inf, -np.inf], np.nan).sort_values(["Year", "Week_Num"])
    return wk

def monthly_totals(df):
    agg = {
        "OrderQty": ("OrderQty", "sum"),
        "CutQty": ("CutQty", "sum"),
        "ShipQty": ("ShipQty", "sum"),
        "CutShipDiff": ("CutShipDiff", "sum"),
    }
    mt = df.groupby(["Year", "Month_Num"], dropna=False, as_index=False).agg(**agg)
    mt["Cut/Ship"] = mt["ShipQty"] / mt["CutQty"]
    mt["Order/Ship"] = mt["ShipQty"] / mt["OrderQty"]
    mt["Order/Cut"] = mt["CutQty"] / mt["OrderQty"]
    mt["Month_Label"] = mt["Month_Num"].map(MONTH_LABEL)
    mt = mt.replace([np.inf, -np.inf], np.nan).sort_values(["Year", "Month_Num"])
    return mt

def top_n_by_ratio(df, group_col, ratio_name, n=10, ascending=False):
    g = df.groupby(group_col, as_index=False).agg(
        OrderQty=("OrderQty", "sum"),
        CutQty=("CutQty", "sum"),
        ShipQty=("ShipQty", "sum"),
    )

    if ratio_name == "Cut/Ship":
        g[ratio_name] = g["ShipQty"] / g["CutQty"]
    elif ratio_name == "Order/Ship":
        g[ratio_name] = g["ShipQty"] / g["OrderQty"]
    elif ratio_name == "Order/Cut":
        g[ratio_name] = g["CutQty"] / g["OrderQty"]
    else:
        raise ValueError("Invalid ratio name")

    g = g.replace([np.inf, -np.inf], np.nan).dropna(subset=[ratio_name])
    g = g.sort_values(ratio_name, ascending=ascending).head(n)
    return g

def diff_breakdown_cols(df):
    possible = [f"Metric {chr(i)}" for i in range(ord("A"), ord("S") + 1)]
    return [c for c in possible if c in df.columns]

def sum_diff_breakdown(df):
    cols = diff_breakdown_cols(df)
    if not cols:
        return pd.Series(dtype=float)
    tmp = df[cols].copy()
    for c in cols:
        tmp[c] = pd.to_numeric(tmp[c], errors="coerce").fillna(0)
    return tmp.sum(numeric_only=True).sort_values(ascending=False)

def build_cutship_from_breakdown(df):
    cols = diff_breakdown_cols(df)
    if not cols:
        return pd.Series([np.nan] * len(df), index=df.index)
    tmp = df[cols].copy()
    for c in cols:
        tmp[c] = pd.to_numeric(tmp[c], errors="coerce").fillna(0)
    return tmp.sum(axis=1)

def top_n_cutshipdiff(df, group_col, n=10):
    g = df.groupby(group_col, as_index=False).agg(
        CutShipDiff=("CutShipDiff", "sum")
    )
    g = g.replace([np.inf, -np.inf], np.nan).dropna(subset=["CutShipDiff"])
    g = g.sort_values("CutShipDiff", ascending=False).head(n)
    return g

def style_plot(fig, percent_y=False, y_title=None, x_title=None, y_range=None):
    fig.update_layout(
        template="plotly_white",
        colorway=CORPORATE_COLORS,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#102a43"),
        title_font=dict(color="#102a43", size=20),
        legend_title_text="",
        margin=dict(l=20, r=20, t=60, b=20)
    )
    fig.update_xaxes(
        showgrid=False,
        title=x_title,
        linecolor="#D9E2EC",
        tickfont=dict(color="#486581")
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="#E9EEF5",
        zeroline=False,
        title=y_title,
        tickfont=dict(color="#486581")
    )
    if percent_y:
        fig.update_yaxes(tickformat=".1%")
    if y_range is not None:
        fig.update_yaxes(range=y_range)
    return fig

def metric_card_html(label, value, delta_text=None):
    delta_html = ""
    if delta_text:
        color = "#2FBF71" if "▲" in delta_text else "#E55353" if "▼" in delta_text else "#486581"
        delta_html = f'<div class="metric-delta" style="color:{color};">{delta_text}</div>'
    return f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {delta_html}
        </div>
    """

def style_ratio_display_table(df_in, ratio_cols):
    df = df_in.copy()
    for c in ratio_cols:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: "NA" if pd.isna(x) else f"{x*100:.1f}%")
    for c in ["OrderQty", "CutQty", "ShipQty", "CutShipDiff", "Qty"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: "NA" if pd.isna(x) else f"{x:,.0f}")
    for c in ["Year", "Week_Num", "Month_Num"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: "NA" if pd.isna(x) else int(x))
    return df

def ratio_tab_chart(df_in, col, ratio_name):
    top = top_n_by_ratio(df_in, col, ratio_name, n=10)
    top_disp = top.copy()
    top_disp[ratio_name] = top_disp[ratio_name] * 100

    ymax = top_disp[ratio_name].max()
    ymax = 101 if pd.isna(ymax) else max(101, float(ymax) + 1)

    fig = px.bar(
        top_disp,
        x=col,
        y=ratio_name,
        text=ratio_name,
        color_discrete_sequence=[CORPORATE_COLORS[0]],
        title=f"Top 10 {col} by {ratio_name}"
    )
    fig.update_traces(
        texttemplate="%{text:.1f}%",
        textposition="outside",
        marker_line_color="white",
        marker_line_width=1.2
    )
    fig = style_plot(
        fig,
        percent_y=False,
        x_title=col,
        y_title=ratio_name,
        y_range=[90, ymax]
    )
    fig.update_yaxes(ticksuffix="%")
    st.plotly_chart(fig, width="stretch")

    tbl = top[[col, "OrderQty", "CutQty", "ShipQty", ratio_name]].copy()
    tbl = style_ratio_display_table(tbl, [ratio_name])
    st.dataframe(tbl, width="stretch", hide_index=True)

def yearly_ratio_summary(df_in):
    yr = df_in.groupby("Year", as_index=False).agg(
        OrderQty=("OrderQty", "sum"),
        CutQty=("CutQty", "sum"),
        ShipQty=("ShipQty", "sum"),
        CutShipDiff=("CutShipDiff", "sum")
    ).sort_values("Year")
    yr["Cut/Ship"] = yr["ShipQty"] / yr["CutQty"]
    yr["Order/Ship"] = yr["ShipQty"] / yr["OrderQty"]
    yr["Order/Cut"] = yr["CutQty"] / yr["OrderQty"]
    yr["CutShipDiff_YoY"] = yr["CutShipDiff"].pct_change()
    return yr

def top5_vertical_ratio_chart(df_in, group_col, ratio_name, color_code):
    top = top_n_by_ratio(df_in, group_col, ratio_name, n=5, ascending=False).copy()
    top = top.sort_values(ratio_name, ascending=False)
    top["RatioPct"] = top[ratio_name] * 100

    ymax = top["RatioPct"].max()
    ymax = 101 if pd.isna(ymax) else max(101, float(ymax) + 1)

    fig = px.bar(
        top,
        x=group_col,
        y="RatioPct",
        text="RatioPct",
        color_discrete_sequence=[color_code],
        title=ratio_name
    )
    fig.update_traces(
        texttemplate="%{text:.1f}%",
        textposition="outside",
        marker_line_color="white",
        marker_line_width=1.2
    )
    fig = style_plot(
        fig,
        percent_y=False,
        x_title=group_col,
        y_title=ratio_name,
        y_range=[85, ymax]
    )
    fig.update_yaxes(ticksuffix="%")
    fig.update_xaxes(tickangle=0)
    return fig

# -----------------------------
# Load data
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    df.columns = [c.replace("  ", " ").strip() for c in df.columns]

    if "Customers" in df.columns:
        df = df.drop(columns=["Customers"])

    df = df.rename(columns={
        "Unit": "Factory",
        "Calling Name": "Customer",
        "Garment item type": "Product"
    })

    df = df.rename(columns={
        "Order Qty": "OrderQty",
        "Cut Qty": "CutQty",
        "Ship Qty": "ShipQty",
        "Cutship Difference": "CutShipDiff",
        "Cut ship Difference": "CutShipDiff",
        "Cut Ship Difference": "CutShipDiff"
    })

    df = df.loc[:, ~df.columns.duplicated()]

    for c in ["Factory", "Customer", "Product", "Week", "Month"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    required = ["Year", "Week", "Factory", "Customer", "Product", "OrderQty", "CutQty", "ShipQty", "Month"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    for c in ["OrderQty", "CutQty", "ShipQty"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    if "CutShipDiff" in df.columns:
        df["CutShipDiff"] = pd.to_numeric(df["CutShipDiff"], errors="coerce")

    if "CutShipDiff" not in df.columns or df["CutShipDiff"].isna().all():
        df["CutShipDiff"] = build_cutship_from_breakdown(df)

    df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
    df["Week_Num"] = clean_week(df["Week"])
    df["Month_Num"] = clean_month(df["Month"])

    df = df.dropna(subset=["OrderQty", "CutQty", "ShipQty"], how="all")
    return df

df = load_data(EXCEL_PATH, SHEET_NAME)

# -----------------------------
# Hero section
# -----------------------------
st.markdown("""
<div class="hero-box">
    <div class="hero-title">Cut to Ship Dashboard</div>
    <div class="hero-sub">Operational performance, ratio movement, cut-ship difference analysis, and year-on-year insights.</div>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# Sidebar filters
# -----------------------------
st.sidebar.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
st.sidebar.header("Filters")

years_all = sorted([int(x) for x in df["Year"].dropna().unique().tolist()])
weeks_all = sorted([int(x) for x in df["Week_Num"].dropna().unique().tolist()])
factories_all = sorted(df["Factory"].dropna().astype(str).str.strip().unique().tolist())
customers_all = sorted(df["Customer"].dropna().astype(str).str.strip().unique().tolist())
products_all = sorted(df["Product"].dropna().astype(str).str.strip().unique().tolist())

f_year = st.sidebar.multiselect("Year", years_all, default=[])
f_week = st.sidebar.multiselect("Week", weeks_all, default=[])
f_factory = st.sidebar.multiselect("Factory", factories_all, default=[])
f_customer = st.sidebar.multiselect("Customer", customers_all, default=[])
f_product = st.sidebar.multiselect("Product", products_all, default=[])

display_years = pick_display_years(df, f_year)

fdf = df.copy()
if display_years:
    fdf = fdf[fdf["Year"].isin(display_years)]
if f_week:
    fdf = fdf[fdf["Week_Num"].isin(f_week)]
if f_factory:
    fdf = fdf[fdf["Factory"].isin(f_factory)]
if f_customer:
    fdf = fdf[fdf["Customer"].isin(f_customer)]
if f_product:
    fdf = fdf[fdf["Product"].isin(f_product)]

fdf_non_year = df.copy()
if f_week:
    fdf_non_year = fdf_non_year[fdf_non_year["Week_Num"].isin(f_week)]
if f_factory:
    fdf_non_year = fdf_non_year[fdf_non_year["Factory"].isin(f_factory)]
if f_customer:
    fdf_non_year = fdf_non_year[fdf_non_year["Customer"].isin(f_customer)]
if f_product:
    fdf_non_year = fdf_non_year[fdf_non_year["Product"].isin(f_product)]

st.sidebar.caption(f"Rows after filters: {len(fdf):,}")
st.sidebar.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page selector
# -----------------------------
page = st.selectbox(
    "Select Page",
    [
        "1 Overall",
        "2 Cut/Ship",
        "3 Order/Ship",
        "4 Order/Cut",
        "5 Cut Ship Difference",
        "6 Cut Ship Difference YOY",
        "7 Latest Week Deep Dive"
    ]
)

# -----------------------------
# Page 1: Overall
# -----------------------------
if page == "1 Overall":
    st.markdown('<div class="page-title">Overall Performance</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">Weekly and monthly ratio movement.</div>', unsafe_allow_html=True)

    total_order = fdf["OrderQty"].sum()
    total_cut = fdf["CutQty"].sum()
    total_ship = fdf["ShipQty"].sum()

    cut_ship = safe_div(total_ship, total_cut)
    order_ship = safe_div(total_ship, total_order)
    order_cut = safe_div(total_cut, total_order)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(metric_card_html("Cut/Ship", percent_fmt(cut_ship, 1)), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card_html("Order/Ship", percent_fmt(order_ship, 1)), unsafe_allow_html=True)
    with c3:
        st.markdown(metric_card_html("Order/Cut", percent_fmt(order_cut, 1)), unsafe_allow_html=True)

    c4, c5, c6 = st.columns(3)
    with c4:
        st.markdown(metric_card_html("Total Order Qty", num_fmt(total_order)), unsafe_allow_html=True)
    with c5:
        st.markdown(metric_card_html("Total Cut Qty", num_fmt(total_cut)), unsafe_allow_html=True)
    with c6:
        st.markdown(metric_card_html("Total Ship Qty", num_fmt(total_ship)), unsafe_allow_html=True)

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Week Wise Ratio Movement")

    wk = weekly_totals(fdf)
    if wk.empty:
        st.warning("No weekly data available after filters.")
    else:
        wk_ratio = wk.melt(
            id_vars=["Year", "Week_Num"],
            value_vars=["Cut/Ship", "Order/Ship", "Order/Cut"],
            var_name="Metric",
            value_name="Value"
        )

        if len(display_years) > 1:
            wk_ratio["Series"] = wk_ratio["Metric"] + " (" + wk_ratio["Year"].astype(str) + ")"
            color_field = "Series"
        else:
            color_field = "Metric"

        max_week = wk["Week_Num"].dropna().max()
        max_week = 52 if pd.isna(max_week) else max(52, int(max_week))

        fig1 = px.line(
            wk_ratio,
            x="Week_Num",
            y="Value",
            color=color_field,
            markers=True,
            color_discrete_sequence=CORPORATE_COLORS
        )
        fig1 = style_plot(fig1, percent_y=True, x_title="Week", y_title="Ratio")
        fig1.update_xaxes(dtick=1, range=[1, max_week])
        st.plotly_chart(fig1, width="stretch")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Month Wise Ratio Movement")

    mt = monthly_totals(fdf)
    if mt.empty:
        st.warning("No monthly data available after filters.")
    else:
        mt_ratio = mt.melt(
            id_vars=["Year", "Month_Num", "Month_Label"],
            value_vars=["Cut/Ship", "Order/Ship", "Order/Cut"],
            var_name="Metric",
            value_name="Value"
        )

        if len(display_years) > 1:
            mt_ratio["Series"] = mt_ratio["Metric"] + " (" + mt_ratio["Year"].astype(str) + ")"
            color_field = "Series"
        else:
            color_field = "Metric"

        fig_month = px.line(
            mt_ratio,
            x="Month_Num",
            y="Value",
            color=color_field,
            markers=True,
            color_discrete_sequence=CORPORATE_COLORS
        )
        fig_month = style_plot(fig_month, percent_y=True, x_title="Month", y_title="Ratio")
        fig_month.update_xaxes(
            tickmode="array",
            tickvals=list(MONTH_LABEL.keys()),
            ticktext=list(MONTH_LABEL.values()),
            range=[1, 12]
        )
        st.plotly_chart(fig_month, width="stretch")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    with st.expander("Weekly totals table", expanded=False):
        wk_disp = wk.copy()
        wk_disp = style_ratio_display_table(
            wk_disp[["Year", "Week_Num", "OrderQty", "CutQty", "ShipQty", "CutShipDiff", "Cut/Ship", "Order/Ship", "Order/Cut"]],
            ["Cut/Ship", "Order/Ship", "Order/Cut"]
        )
        st.dataframe(wk_disp, width="stretch", hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page 2: Cut/Ship
# -----------------------------
elif page == "2 Cut/Ship":
    st.markdown('<div class="page-title">Cut/Ship</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Customer (Top 10)", "Product (Top 10)", "Factory (Top 10)"])

    with tabs[0]:
        ratio_tab_chart(fdf, "Customer", "Cut/Ship")

    with tabs[1]:
        ratio_tab_chart(fdf, "Product", "Cut/Ship")

    with tabs[2]:
        ratio_tab_chart(fdf, "Factory", "Cut/Ship")

# -----------------------------
# Page 3: Order/Ship
# -----------------------------
elif page == "3 Order/Ship":
    st.markdown('<div class="page-title">Order/Ship</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Customer (Top 10)", "Product (Top 10)", "Factory (Top 10)"])

    with tabs[0]:
        ratio_tab_chart(fdf, "Customer", "Order/Ship")

    with tabs[1]:
        ratio_tab_chart(fdf, "Product", "Order/Ship")

    with tabs[2]:
        ratio_tab_chart(fdf, "Factory", "Order/Ship")

# -----------------------------
# Page 4: Order/Cut
# -----------------------------
elif page == "4 Order/Cut":
    st.markdown('<div class="page-title">Order/Cut</div>', unsafe_allow_html=True)
    tabs = st.tabs(["Customer (Top 10)", "Product (Top 10)", "Factory (Top 10)"])

    with tabs[0]:
        ratio_tab_chart(fdf, "Customer", "Order/Cut")

    with tabs[1]:
        ratio_tab_chart(fdf, "Product", "Order/Cut")

    with tabs[2]:
        ratio_tab_chart(fdf, "Factory", "Order/Cut")

# -----------------------------
# Page 5: Cut Ship Difference
# -----------------------------
elif page == "5 Cut Ship Difference":
    st.markdown('<div class="page-title">Cut Ship Difference Analysis</div>', unsafe_allow_html=True)

    if "CutShipDiff" not in fdf.columns:
        st.error("Column 'Cut Ship Difference' was not found or created.")
    else:
        total_diff = fdf["CutShipDiff"].sum()
        total_order = fdf["OrderQty"].sum()
        total_cut = fdf["CutQty"].sum()
        total_ship = fdf["ShipQty"].sum()

        pct_order = safe_div(total_diff, total_order)
        pct_cut = safe_div(total_diff, total_cut)
        pct_ship = safe_div(total_diff, total_ship)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(metric_card_html("Total Cut Ship Difference", num_fmt(total_diff)), unsafe_allow_html=True)
        with c2:
            st.markdown(metric_card_html("As % of Order Qty", percent_fmt(pct_order, 1)), unsafe_allow_html=True)
        with c3:
            st.markdown(metric_card_html("As % of Cut Qty", percent_fmt(pct_cut, 1)), unsafe_allow_html=True)
        with c4:
            st.markdown(metric_card_html("As % of Ship Qty", percent_fmt(pct_ship, 1)), unsafe_allow_html=True)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.subheader("Reason Wise Breakdown of Cut Ship Difference")

        breakdown = sum_diff_breakdown(fdf)
        if breakdown.empty:
            st.warning("No breakdown columns found from Metric A to Metric S.")
        else:
            bdf = breakdown.reset_index()
            bdf.columns = ["Reason", "Qty"]

            fig_reason = px.pie(
                bdf,
                names="Reason",
                values="Qty",
                color_discrete_sequence=CORPORATE_COLORS
            )
            fig_reason.update_traces(textinfo="percent+label")
            fig_reason.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig_reason, width="stretch")

            bdf = style_ratio_display_table(bdf, [])
            st.dataframe(bdf, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.subheader("Factory, Customer and Product Wise Cut Ship Difference")

        t1, t2, t3 = st.tabs(["Factory", "Customer (Top 10)", "Product (Top 10)"])

        with t1:
            fac = fdf.groupby("Factory", as_index=False)["CutShipDiff"].sum().sort_values("CutShipDiff", ascending=False)
            fig_fac = px.pie(fac, names="Factory", values="CutShipDiff", color_discrete_sequence=CORPORATE_COLORS)
            fig_fac.update_traces(textinfo="percent+label")
            fig_fac.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig_fac, width="stretch")
            fac = style_ratio_display_table(fac, [])
            st.dataframe(fac, width="stretch", hide_index=True)

        with t2:
            cust = top_n_cutshipdiff(fdf, "Customer", 10)
            fig_cust = px.pie(cust, names="Customer", values="CutShipDiff", color_discrete_sequence=CORPORATE_COLORS)
            fig_cust.update_traces(textinfo="percent+label")
            fig_cust.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig_cust, width="stretch")
            cust = style_ratio_display_table(cust, [])
            st.dataframe(cust, width="stretch", hide_index=True)

        with t3:
            prod = top_n_cutshipdiff(fdf, "Product", 10)
            fig_prod = px.pie(prod, names="Product", values="CutShipDiff", color_discrete_sequence=CORPORATE_COLORS)
            fig_prod.update_traces(textinfo="percent+label")
            fig_prod.update_layout(template="plotly_white", margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig_prod, width="stretch")
            prod = style_ratio_display_table(prod, [])
            st.dataframe(prod, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page 6: YOY Cut Ship Difference
# -----------------------------
elif page == "6 Cut Ship Difference YOY":
    st.markdown('<div class="page-title">Year on Year Cut Ship Difference</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">Year summary with reason movement comparison.</div>', unsafe_allow_html=True)

    yoy = yearly_ratio_summary(fdf_non_year)
    yoy_show = yoy[yoy["Year"].isin([2024, 2025, 2026])].copy().sort_values("Year")

    if yoy_show.empty:
        st.warning("No year-on-year data available for 2024 to 2026.")
    else:
        cols = st.columns(len(yoy_show))
        for i, (_, row) in enumerate(yoy_show.iterrows()):
            prev_val = yoy_show.iloc[i - 1]["CutShipDiff"] if i > 0 else np.nan
            with cols[i]:
                st.markdown(
                    metric_card_html(
                        f"{int(row['Year'])} Cut Ship Difference",
                        num_fmt(row["CutShipDiff"]),
                        delta_fmt(row["CutShipDiff"], prev_val, 1)
                    ),
                    unsafe_allow_html=True
                )

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Year on Year Breakdown by Reason")

    bcols = diff_breakdown_cols(fdf_non_year)
    if bcols:
        tmp = fdf_non_year.copy()
        for c in bcols:
            tmp[c] = pd.to_numeric(tmp[c], errors="coerce").fillna(0)

        yoy_break = tmp.groupby("Year", as_index=False)[bcols].sum()
        yoy_break = yoy_break[yoy_break["Year"].isin([2024, 2025, 2026])].sort_values("Year")

        yoy_long = yoy_break.melt(
            id_vars=["Year"],
            value_vars=bcols,
            var_name="Reason",
            value_name="Qty"
        )

        top_reasons = (
            yoy_long.groupby("Reason", as_index=False)["Qty"]
            .sum()
            .sort_values("Qty", ascending=False)
            .head(12)["Reason"]
            .tolist()
        )
        yoy_long = yoy_long[yoy_long["Reason"].isin(top_reasons)]

        fig_reason_yoy = px.bar(
            yoy_long,
            x="Year",
            y="Qty",
            color="Reason",
            barmode="relative",
            color_discrete_sequence=CORPORATE_COLORS
        )
        fig_reason_yoy = style_plot(fig_reason_yoy, percent_y=False, x_title="Year", y_title="Qty")
        fig_reason_yoy.update_traces(
            hovertemplate="Year=%{x}<br>Reason=%{fullData.name}<br>Qty=%{y:,.0f}<extra></extra>"
        )
        st.plotly_chart(fig_reason_yoy, width="stretch")

        yoy_break = style_ratio_display_table(yoy_break, [])
        st.dataframe(yoy_break, width="stretch", hide_index=True)
    else:
        st.warning("No breakdown columns found from Metric A to Metric S.")
    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page 7: Latest Week Deep Dive
# -----------------------------
elif page == "7 Latest Week Deep Dive":
    st.markdown('<div class="page-title">Latest Week Deep Dive</div>', unsafe_allow_html=True)

    latest_year = int(df["Year"].dropna().max())
    latest_week = int(df[df["Year"] == latest_year]["Week_Num"].dropna().max())
    st.info(f"Latest Year: {latest_year} | Latest Week: {latest_week}")

    base = df[(df["Year"] == latest_year) & (df["Week_Num"] == latest_week)].copy()

    if f_factory:
        base = base[base["Factory"].isin(f_factory)]
    if f_customer:
        base = base[base["Customer"].isin(f_customer)]
    if f_product:
        base = base[base["Product"].isin(f_product)]

    if base.empty:
        st.warning("No data available for latest week after applying filters.")
    else:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.subheader("Factory wise summary (latest week)")
        fac = base.groupby("Factory", as_index=False).agg(
            OrderQty=("OrderQty", "sum"),
            CutQty=("CutQty", "sum"),
            ShipQty=("ShipQty", "sum"),
            CutShipDiff=("CutShipDiff", "sum")
        )
        fac["Cut/Ship"] = fac["ShipQty"] / fac["CutQty"]
        fac["Order/Ship"] = fac["ShipQty"] / fac["OrderQty"]
        fac["Diff/ShipQty"] = fac["CutShipDiff"] / fac["ShipQty"]
        fac = fac.replace([np.inf, -np.inf], np.nan)

        fac = style_ratio_display_table(fac.sort_values("Cut/Ship"), ["Cut/Ship", "Order/Ship", "Diff/ShipQty"])
        st.dataframe(fac, width="stretch", hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.subheader("Selecting one Factory for deeper diagnosis")

        sel_factory = st.selectbox("Factory", sorted(base["Factory"].dropna().astype(str).unique().tolist()))
        b1 = base[base["Factory"] == sel_factory].copy()

        cust = b1.groupby("Customer", as_index=False).agg(
            OrderQty=("OrderQty", "sum"),
            CutQty=("CutQty", "sum"),
            ShipQty=("ShipQty", "sum")
        )
        cust["Cut/Ship"] = cust["ShipQty"] / cust["CutQty"]
        cust["Order/Ship"] = cust["ShipQty"] / cust["OrderQty"]
        cust["Order/Cut"] = cust["CutQty"] / cust["OrderQty"]
        cust = cust.replace([np.inf, -np.inf], np.nan)

        prod = b1.groupby("Product", as_index=False).agg(
            OrderQty=("OrderQty", "sum"),
            CutQty=("CutQty", "sum"),
            ShipQty=("ShipQty", "sum")
        )
        prod["Cut/Ship"] = prod["ShipQty"] / prod["CutQty"]
        prod["Order/Ship"] = prod["ShipQty"] / prod["OrderQty"]
        prod["Order/Cut"] = prod["CutQty"] / prod["OrderQty"]
        prod = prod.replace([np.inf, -np.inf], np.nan)

        t1, t2 = st.tabs(["Customer drivers", "Product drivers"])

        with t1:
            st.plotly_chart(top5_vertical_ratio_chart(cust, "Customer", "Cut/Ship", CORPORATE_COLORS[0]), width="stretch")
            st.plotly_chart(top5_vertical_ratio_chart(cust, "Customer", "Order/Ship", CORPORATE_COLORS[2]), width="stretch")
            st.plotly_chart(top5_vertical_ratio_chart(cust, "Customer", "Order/Cut", CORPORATE_COLORS[3]), width="stretch")

            cust_table = style_ratio_display_table(cust, ["Cut/Ship", "Order/Ship", "Order/Cut"])
            st.dataframe(cust_table, width="stretch", hide_index=True)

        with t2:
            st.plotly_chart(top5_vertical_ratio_chart(prod, "Product", "Cut/Ship", CORPORATE_COLORS[0]), width="stretch")
            st.plotly_chart(top5_vertical_ratio_chart(prod, "Product", "Order/Ship", CORPORATE_COLORS[2]), width="stretch")
            st.plotly_chart(top5_vertical_ratio_chart(prod, "Product", "Order/Cut", CORPORATE_COLORS[3]), width="stretch")

            prod_table = style_ratio_display_table(prod, ["Cut/Ship", "Order/Ship", "Order/Cut"])
            st.dataframe(prod_table, width="stretch", hide_index=True)

        st.markdown('</div>', unsafe_allow_html=True)
'''

with open(APP_FILE, "w", encoding="utf-8") as f:
    f.write(textwrap.dedent(app_code).lstrip())

print(f"Wrote Streamlit app to: {os.path.abspath(APP_FILE)}")
print(f"Running Streamlit on port {PORT} ...")
print(f"Open in browser: http://localhost:{PORT}")

cmd = [sys.executable, "-m", "streamlit", "run", APP_FILE, "--server.port", str(PORT)]
subprocess.run(cmd)