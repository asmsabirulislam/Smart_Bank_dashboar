"""
Bank Submit History — Streamlit Dashboard
==========================================
File: bank_dashboard_generator.py

Setup (once):
    pip install streamlit pandas plotly openpyxl reportlab

Run:
    streamlit run bank_dashboard_generator.py
"""

import os, glob, shutil, tempfile
from datetime import datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from io import BytesIO

try:
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, PageBreak
    from reportlab.pdfbase import pdfmetrics
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Bank Submit Dashboard", page_icon="🏦",
                   layout="wide", initial_sidebar_state="expanded")

# ─────────────────────────────────────────────────────────────────────────────
# THEME TOGGLE
# ─────────────────────────────────────────────────────────────────────────────
def _apply_theme_css(theme: str):
    if theme == "light":
        bg = "#f6f7fb"; sidebar_bg = "#ffffff"; text = "#0b1220"
        mut = "#445066"; border = "rgba(11,18,32,0.15)"
        btn_bg = "#00c9a7"; btn_fg = "#071824"; btn_bg_hover = "#00b397"
        input_bg = "#ffffff"
    else:
        bg = "#0b1220"; sidebar_bg = "#1a1f2c"; text = "#e2eaf3"
        mut = "#8899aa"; border = "rgba(226,234,243,0.16)"
        btn_bg = "#00c9a7"; btn_fg = "#071824"; btn_bg_hover = "#00b397"
        input_bg = "#0f1626"

    st.markdown(f"""
    <style>
    .stApp {{ background: {bg}; color: {text}; }}
    section[data-testid='stSidebar'] > div:first-child {{ background: {sidebar_bg}; }}
    textarea, input, select {{
        background: {input_bg} !important;
        color: {text} !important;
        border-color: {border} !important;
    }}
    button[kind='primary'],
    button[data-testid='baseButton-primary'] {{
        background-color: {btn_bg} !important;
        color: {btn_fg} !important;
        border: 1px solid rgba(0,0,0,0.15) !important;
    }}
    button[kind='primary']:hover,
    button[data-testid='baseButton-primary']:hover {{
        background-color: {btn_bg_hover} !important;
    }}
    button[kind='secondary'],
    button[data-testid='baseButton-secondary'] {{
        background-color: transparent !important;
        color: {text} !important;
        border: 1px solid {border} !important;
    }}
    p.sh {{ color:#8899aa; font-size:13px; font-weight:600;
             margin-bottom:4px; letter-spacing:.05em; }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — Theme first, then file upload, then filters
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 🏦 Bank Submit\nDashboard")
theme_choice = st.sidebar.radio("Theme", ["dark", "light"], index=0)
_apply_theme_css(theme_choice)
st.sidebar.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# PLOTLY THEME
# ─────────────────────────────────────────────────────────────────────────────
PL = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#8899aa", size=11, family="monospace"),
    margin=dict(l=8, r=8, t=32, b=8),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=10)),
    xaxis=dict(gridcolor="#1a2a3a", linecolor="#1a2a3a"),
    yaxis=dict(gridcolor="#1a2a3a", linecolor="#1a2a3a"),
)
PL_GENERAL = {k: v for k, v in PL.items() if k not in ("legend", "xaxis", "yaxis")}

C = ["#00c9a7","#1a8fff","#ff6b35","#ffd700","#cc44ff",
     "#ff3366","#44aaff","#ff9944","#66dd66","#00aaff"]

REQUIRED_COLUMNS = [
    "Firm Name", "Sales Person", "Bank Submition Date", "Invoice Value",
    "Lc Value", "Maturity Date", "Payment. Rcv Dt", "Bank Accept Date",
    "LC No", "Our Bank", "Party Name", "Bank Name"
]

def usd(v):
    try: v = float(v)
    except Exception: return "$0.00"
    if v >= 1e6: return f"${v/1e6:.2f}M"
    if v >= 1e3: return f"${v/1e3:.1f}K"
    return f"${v:.2f}"

def sh(label):
    st.markdown(f'<p class="sh">{label}</p>', unsafe_allow_html=True)

def norm_tenor(t):
    if pd.isna(t): return "Unknown"
    tt = str(t).strip()
    if tt.startswith("120"): return "120 Days"
    if tt.startswith("90"):  return "90 Days"
    if tt == "0" or "at sight" in tt.lower(): return "At Sight"
    return tt

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADER
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="⏳ Loading data…")
def load(path):
    xls = pd.ExcelFile(path)
    sheet_candidates = ["Raw Data", "raw data", "Sheet1", "Sheet 1"]
    selected = None
    for s in sheet_candidates:
        if s in xls.sheet_names:
            selected = s; break
    if selected is None:
        for s in xls.sheet_names:
            if "bank" in s.lower() and "history" in s.lower():
                selected = s; break
    if selected is None:
        selected = xls.sheet_names[0]

    df = pd.read_excel(path, sheet_name=selected)
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet '{selected}' — Missing columns: {', '.join(missing)}")

    df = df.dropna(subset=["Firm Name"])
    df["Sales Person"] = (df["Sales Person"].astype(str).str.strip()
                          .str.replace("_x000D_\n", "", regex=False).str.strip())
    df["Sales Person"] = df["Sales Person"].replace(r"^(nan|NaN|none|None|\s*)$", None, regex=True)
    df.loc[df["Sales Person"].isin([None, ""]), "Sales Person"] = None

    df["_date"]     = pd.to_datetime(df["Bank Submition Date"], errors="coerce")
    df["MonthSort"] = df["_date"].dt.to_period("M")
    df["Month"]     = df["_date"].dt.strftime("%b %Y")
    df["WeekSort"]  = df["_date"].dt.to_period("W")
    df["Week"]      = df["_date"].dt.strftime("W%V") + " " + df["_date"].dt.strftime("%b %Y")
    df["DayName"]   = df["_date"].dt.strftime("%a")
    df["Date"]      = df["_date"].dt.strftime("%d %b %Y")
    return df

# ─────────────────────────────────────────────────────────────────────────────
# FILE PICKER
# ─────────────────────────────────────────────────────────────────────────────
xlsx = [f for f in glob.glob("*.xlsx") + glob.glob("**/*.xlsx", recursive=True)
        if "Dashboard" not in f]
up = st.sidebar.file_uploader("📂 Upload Excel file", type=["xlsx"])
if up:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    shutil.copyfileobj(up, tmp); tmp.close(); FP = tmp.name
elif xlsx:
    FP = st.sidebar.selectbox("Or pick a file", xlsx)
else:
    st.warning("⚠️ Please upload your Excel file using the sidebar."); st.stop()

raw = load(FP)

# ─────────────────────────────────────────────────────────────────────────────
# FILTERS
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.markdown("### 🔽 Filters")
ml  = [str(m) for m in sorted(raw["MonthSort"].dropna().unique())]
sm  = st.sidebar.multiselect("Month",    ml, default=ml)
sf  = st.sidebar.multiselect("Firm",     sorted(raw["Firm Name"].dropna().unique()),
                              default=sorted(raw["Firm Name"].dropna().unique()))
sb  = st.sidebar.multiselect("Our Bank", sorted(raw["Our Bank"].dropna().unique()),
                              default=sorted(raw["Our Bank"].dropna().unique()))

sales_persons = sorted(raw["Sales Person"].dropna().unique())
ss_choices    = (["(Blank)"] + sales_persons) if raw["Sales Person"].isna().any() else sales_persons
ss = st.sidebar.multiselect("Sales Person", ss_choices, default=ss_choices)

sparty = st.sidebar.multiselect("Party Name",
    sorted(raw["Party Name"].dropna().unique()),
    default=sorted(raw["Party Name"].dropna().unique()))

min_date = raw["_date"].min(); max_date = raw["_date"].max()
if pd.isna(min_date) or pd.isna(max_date):
    date_range = st.sidebar.date_input("Date Range",
        value=(pd.Timestamp.today().date(), pd.Timestamp.today().date()))
else:
    date_range = st.sidebar.date_input("Date Range",
        value=(min_date.date(), max_date.date()))

df = raw.copy()
if sm:     df = df[df["MonthSort"].astype(str).isin(sm)]
if sf:     df = df[df["Firm Name"].isin(sf)]
if sb:     df = df[df["Our Bank"].isin(sb)]
if ss:
    if "(Blank)" in ss:
        sel = [s for s in ss if s != "(Blank)"]
        df = df[(df["Sales Person"].isin(sel)) | df["Sales Person"].isna()]
    else:
        df = df[df["Sales Person"].isin(ss)]
if sparty: df = df[df["Party Name"].isin(sparty)]
if isinstance(date_range, tuple) and len(date_range) == 2:
    s_d, e_d = date_range
    df = df[(df["_date"].dt.date >= s_d) & (df["_date"].dt.date <= e_d)]

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(df):,}** of **{len(raw):,}** records")
if df.empty: st.warning("No records match the filters."); st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# AGGREGATES
# ─────────────────────────────────────────────────────────────────────────────
N      = len(df)
inv    = df["Invoice Value"].sum()
lc     = df["Lc Value"].sum()
mat_v  = df[df["Maturity Date"].notna()]["Invoice Value"].sum()
pay_v  = df[df["Payment. Rcv Dt"].notna()]["Invoice Value"].sum()
paid_n = int(df["Payment. Rcv Dt"].notna().sum())
acc_n  = int((df["Bank Accept Date"].notna() & df["Payment. Rcv Dt"].isna()).sum())
nacc_n = int(df["Bank Accept Date"].isna().sum())
acc_v  = df[df["Bank Accept Date"].notna() & df["Payment. Rcv Dt"].isna()]["Invoice Value"].sum()
nacc_v = df[df["Bank Accept Date"].isna()]["Invoice Value"].sum()

monthly = (df.groupby(["MonthSort","Month"])
             .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"), LC=("Lc Value","sum"))
             .reset_index().sort_values("MonthSort"))
by_firm = (df.groupby("Firm Name")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False))
by_bank = (df.groupby("Our Bank")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False))
t_party = (df.groupby("Party Name")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False).head(10))
t_bname = (df.groupby("Bank Name")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False).head(10))

spg   = (df[df["Sales Person"].notna()]
          .groupby("Sales Person")
          .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
          .reset_index().sort_values("Inv", ascending=False))
sp_p  = (df[df["Payment. Rcv Dt"].notna() & df["Sales Person"].notna()]
          .groupby("Sales Person").size().reset_index(name="Paid"))
spg   = spg.merge(sp_p, on="Sales Person", how="left").fillna(0)
spg["Pct"] = (spg["Paid"] / spg["N"] * 100).round(1)

# Weekly
weekly = (df.groupby(["WeekSort","Week"])
           .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"), LC=("Lc Value","sum"),
                Paid_n=("Payment. Rcv Dt", lambda x: x.notna().sum()))
           .reset_index().sort_values("WeekSort"))
weekly["Paid_pct"] = (weekly["Paid_n"] / weekly["Count"] * 100).round(1)

wk_firm = (df.groupby(["WeekSort","Week","Firm Name"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))
wk_sp   = (df[df["Sales Person"].notna()]
             .groupby(["WeekSort","Week","Sales Person"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))
wk_bank = (df.groupby(["WeekSort","Week","Our Bank"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))

wk_status = df.copy()
wk_status["Status"] = wk_status.apply(lambda r:
    "Paid"         if pd.notna(r["Payment. Rcv Dt"]) else
    "Accepted"     if pd.notna(r["Bank Accept Date"]) else
    "Not Accepted", axis=1)
wk_st_grp = (wk_status.groupby(["WeekSort","Week","Status"])
              .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"))
              .reset_index().sort_values("WeekSort"))

wk_party     = (df.groupby(["WeekSort","Week","Party Name"])
                  .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
                  .reset_index().sort_values("WeekSort"))
wk_party_top = wk_party[wk_party["Party Name"].isin(t_party["Party Name"].tolist())]

period = (f"{monthly['Month'].iloc[0]} – {monthly['Month'].iloc[-1]}"
          if len(monthly) else "—")

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("# 🏦 Bank Submit History Dashboard")
st.markdown(
    f"<p style='color:#556677;font-size:12px;margin-top:-12px;'>"
    f"Period: <b style='color:#00c9a7'>{period}</b> &nbsp;|&nbsp; "
    f"File: {os.path.basename(FP)}</p>", unsafe_allow_html=True)
st.markdown("---")

c1, c2, c3, c4 = st.columns(4)
c1.metric("📋 Total Submissions",       f"{N:,}")
c2.metric("💵 Total Invoice Value",     usd(inv))
c3.metric("📅 Maturity Received Value", usd(mat_v))
c4.metric("✅ Payment Received Value",  usd(pay_v),
          delta=f"{paid_n} records · {paid_n/N*100:.1f}%")
st.markdown("")

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
(t_daily, t_overview, t_weekly, t_firm, t_banks, t_parties, t_payment) = st.tabs([
    "📅 Daily Analysis",
    "📊 Overview",
    "📅 Weekly Analysis",
    "🏢 Firm & Sales Person",
    "🏦 Banks",
    "👥 Top Parties",
    "🔄 Payment Status",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — DAILY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t_daily:
    # ── Selected-date selector (today by default) ─────────────────────────
    min_d = df['_date'].dt.date.min()
    max_d = df['_date'].dt.date.max()
    today_d = pd.Timestamp.today().date()
    default_sel = today_d if (min_d is not pd.NaT and max_d is not pd.NaT and min_d <= today_d <= max_d) else (max_d if pd.notna(max_d) else today_d)

    daily_sel = st.date_input(
        "📅 Select Date for Daily Analysis",
        value=default_sel,
        min_value=min_d if pd.notna(min_d) else None,
        max_value=max_d if pd.notna(max_d) else None,
        key="daily_sel_date",
    )

    df_daily = df[df["_date"].dt.normalize().dt.date == daily_sel].copy()
    daily_N = len(df_daily)

    st.caption(f"Daily data for: {pd.to_datetime(daily_sel).strftime('%d %b %Y')}  |  Records: {daily_N:,}")

    if df_daily.empty:
        st.warning("No records for the selected date (after applying global filters).")
        st.stop()

    # ── KPI cards ─────────────────────────────────────────────────────────
    daily_qty   = df_daily["Invoice Qty"].sum() if "Invoice Qty" in df_daily.columns else 0
    daily_avg   = df_daily["Invoice Value"].sum() / daily_N if daily_N else 0
    daily_inv   = df_daily["Invoice Value"].sum() if "Invoice Value" in df_daily.columns else 0
    paid_amt    = df_daily[df_daily["Payment. Rcv Dt"].notna()]["Invoice Value"].sum()
    pending_amt = df_daily[df_daily["Payment. Rcv Dt"].isna()]["Invoice Value"].sum()

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("📋 Total Submissions", f"{daily_N:,}")
    k2.metric("💵 Invoice Value",     usd(daily_inv))
    k3.metric("📦 Invoice Qty",       f"{daily_qty:,.0f}")
    k4.metric("📈 Avg Value",         usd(daily_avg))
    k5.metric("🏢 Unique Parties",    f"{df_daily['Party Name'].nunique():,}")
    k6.metric("🏦 Unique Banks",      f"{df_daily['Our Bank'].nunique():,}")

    st.markdown("")
    r1, r2 = st.columns(2)
    r1.metric("✅ Accepted Invoice Value", usd(paid_amt))
    r2.metric("⏳ Pending Invoice Value",  usd(pending_amt))
    st.markdown("---")

    # ── Daily table + charts ──────────────────────────────────────────────
    daily_grp = (df.groupby(df["_date"].dt.date)
                   .agg(Submissions=("LC No","count"),
                        Qty=("Invoice Qty","sum"),
                        Value=("Invoice Value","sum"))
                   .reset_index().rename(columns={"_date":"Date"}))
    if not daily_grp.empty:
        daily_grp["Date"] = pd.to_datetime(daily_grp["Date"])
        daily_grp = daily_grp.sort_values("Date")
        daily_grp["Cumulative Value"] = daily_grp["Value"].cumsum()

        dc1, dc2, dc3 = st.columns([2, 2, 3])
        with dc1:
            sh("📅 Daily Summary Table")
            tbl = daily_grp.copy()
            tbl["Date"]             = tbl["Date"].dt.strftime("%d %b %Y")
            tbl["Value"]            = tbl["Value"].map(lambda x: f"${x:,.2f}")
            tbl["Cumulative Value"] = tbl["Cumulative Value"].map(lambda x: f"${x:,.2f}")
            tbl["Qty"]              = tbl["Qty"].map(lambda x: f"{x:,.0f}")
            st.dataframe(tbl, use_container_width=True, hide_index=True, height=360)

        with dc2:
            sh("📈 Daily Invoice Value")
            fig_d = go.Figure()
            fig_d.add_bar(x=daily_grp["Date"], y=daily_grp["Value"], marker_color="#1a8fff")
            fig_d.update_layout(**PL_GENERAL,
                xaxis=dict(title="Date", tickangle=-40, tickfont=dict(size=9), gridcolor="#1a2a3a"),
                yaxis=dict(title="Invoice Value (USD)", gridcolor="#1a2a3a"),
                height=320, showlegend=False)
            st.plotly_chart(fig_d, use_container_width=True)

        with dc3:
            sh("📈 Cumulative Invoice Value")
            fig_cum = go.Figure()
            fig_cum.add_scatter(x=daily_grp["Date"], y=daily_grp["Cumulative Value"],
                                mode="lines+markers",
                                line=dict(color="#00c9a7", width=2.5),
                                fill="tozeroy", fillcolor="rgba(0,201,167,0.08)")
            fig_cum.update_layout(**PL_GENERAL,
                xaxis=dict(title="Date", tickangle=-40, tickfont=dict(size=9), gridcolor="#1a2a3a"),
                yaxis=dict(title="Cumulative Value (USD)", gridcolor="#1a2a3a"),
                height=320, showlegend=False)
            st.plotly_chart(fig_cum, use_container_width=True)
    else:
        st.warning("No daily data available for the current filters.")

    st.markdown("---")

    # ── Our Bank Breakdown ────────────────────────────────────────────────
    banks_order  = ["SEBPLC","PBL","CBP","DBBL","ONE"]
    by_bank_full = (df_daily.groupby("Our Bank")
                      .agg(Value=("Invoice Value","sum"), Submissions=("LC No","count"))
                      .reset_index().sort_values("Value", ascending=False))
    parts = [by_bank_full[by_bank_full["Our Bank"] == b]
             for b in banks_order if b in by_bank_full["Our Bank"].values]
    rest  = by_bank_full[~by_bank_full["Our Bank"].isin(banks_order)]
    if not rest.empty: parts.append(rest)
    by_bank_ord = pd.concat(parts, ignore_index=True) if parts else by_bank_full

    br1, br2 = st.columns(2)
    with br1:
        sh("🏦 Our Bank Breakdown")
        fig_pie = px.pie(by_bank_ord, names="Our Bank", values="Value",
                         hole=0.48, color_discrete_sequence=C)
        fig_pie.update_layout(**PL)
        fig_pie.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig_pie, use_container_width=True)
    with br2:
        sh("🏦 Our Bank Table")
        tb = by_bank_ord.copy()
        tb["Value"] = tb["Value"].map(lambda x: f"${x:,.2f}")
        tb.columns  = ["Our Bank","Invoice Value","Submissions"]
        st.dataframe(tb, use_container_width=True, hide_index=True, height=320)

    st.markdown("---")

    # ── Firm Name Breakdown ───────────────────────────────────────────────
    by_firm_full = (df_daily.groupby("Firm Name")
                      .agg(Value=("Invoice Value","sum"), Submissions=("LC No","count"))
                      .reset_index().sort_values("Value", ascending=False))
    top_firms = by_firm_full.head(10)
    f1, f2 = st.columns(2)
    with f1:
        sh("🏢 Firm Name Breakdown")
        fig_f = px.bar(top_firms, y="Firm Name", x="Value", orientation="h",
                       color="Firm Name", color_discrete_sequence=C,
                       text=top_firms["Value"].map(usd))
        fig_f.update_traces(textposition="outside", textfont_size=10)
        fig_f.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=360)
        st.plotly_chart(fig_f, use_container_width=True)
    with f2:
        sh("🏢 Firm Name Table")
        tf = top_firms.copy()
        tf["Value"] = tf["Value"].map(lambda x: f"${x:,.2f}")
        tf.columns  = ["Firm Name","Invoice Value","Submissions"]
        st.dataframe(tf, use_container_width=True, hide_index=True, height=360)

    st.markdown("---")

    # ── Sales Person Performance ──────────────────────────────────────────
    spg_d = (df_daily[df_daily["Sales Person"].notna()]
               .groupby("Sales Person")
               .agg(Value=("Invoice Value","sum"), N=("LC No","count"))
               .reset_index().sort_values("Value", ascending=False))
    if not spg_d.empty:
        sp_paid = (df_daily[df_daily["Payment. Rcv Dt"].notna() & df_daily["Sales Person"].notna()]
                     .groupby("Sales Person").size().reset_index(name="Paid"))
        spg_d = spg_d.merge(sp_paid, on="Sales Person", how="left").fillna(0)
        spg_d["Pct"] = (spg_d["Paid"] / spg_d["N"] * 100).round(1)
        total_val_d  = spg_d["Value"].sum()
        spg_d["% of Total"] = (spg_d["Value"] / total_val_d * 100).round(1).map(lambda x: f"{x:.1f}%")

    s1, s2 = st.columns(2)
    with s1:
        sh("👤 Sales Person Performance")
        sp_show = spg_d[["Sales Person","Value","N","Paid","Pct","% of Total"]].copy()
        sp_show["Value"] = sp_show["Value"].map(lambda x: f"${x:,.2f}")
        sp_show["Pct"]   = sp_show["Pct"].map(lambda x: f"{x:.1f}%")
        sp_show["Paid"]  = sp_show["Paid"].astype(int)
        sp_show.columns  = ["Sales Person","Invoice Value","Submissions","Paid","Pay Rate","% of Total"]
        st.dataframe(sp_show, use_container_width=True, hide_index=True, height=360)
    with s2:
        sh("👤 Sales Person Chart")
        fig_sp = px.bar(spg_d.head(12), x="Value", y="Sales Person", orientation="h",
                        color="Sales Person", color_discrete_sequence=C,
                        text=spg_d.head(12)["Value"].map(usd))
        fig_sp.update_traces(textposition="outside", textfont_size=10)
        fig_sp.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=360)
        st.plotly_chart(fig_sp, use_container_width=True)

    st.markdown("---")

    # ── Tenor Distribution ────────────────────────────────────────────────
    ten_dist = pd.DataFrame()
    if "Tenor" in df_daily.columns:
        ten = df_daily["Tenor"].apply(norm_tenor)
        ten_dist = ten.value_counts().reset_index()
        ten_dist.columns = ["Tenor","Count"]

    t1a, t1b = st.columns(2)
    with t1a:
        sh("⏱ Tenor Distribution")
        if not ten_dist.empty:
            fig_t = px.pie(ten_dist, names="Tenor", values="Count",
                           color_discrete_sequence=C, hole=0.45)
            fig_t.update_layout(**PL)
            fig_t.update_traces(textinfo="label+percent", textfont_size=11)
            st.plotly_chart(fig_t, use_container_width=True)
        else:
            st.warning("Tenor column not available.")
    with t1b:
        sh("⏱ Tenor Distribution Table")
        if not ten_dist.empty:
            st.dataframe(ten_dist, use_container_width=True, hide_index=True, height=300)

    st.markdown("---")

    # ── Top 10 Party Names ────────────────────────────────────────────────
    t_party_d = (df_daily.groupby("Party Name")
                   .agg(Value=("Invoice Value","sum"), N=("LC No","count"))
                   .reset_index().sort_values("Value", ascending=False).head(10))
    p1, p2 = st.columns(2)
    with p1:
        sh("🏭 Top 10 Party Names")
        fig_tp = px.bar(t_party_d.sort_values("Value"), x="Value", y="Party Name",
                        orientation="h", color="Party Name", color_discrete_sequence=C,
                        text=t_party_d["Value"].map(usd))
        fig_tp.update_traces(textposition="outside", textfont_size=9)
        fig_tp.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False, height=360)
        st.plotly_chart(fig_tp, use_container_width=True)
    with p2:
        sh("🏭 Top 10 Party Names Table")
        tbp = t_party_d.copy()
        tbp.insert(0, "Rank", range(1, len(tbp)+1))
        tbp["Value"] = tbp["Value"].map(lambda x: f"${x:,.2f}")
        tbp.columns  = ["Rank","Party Name","Invoice Value","Submissions"]
        st.dataframe(tbp, use_container_width=True, hide_index=True, height=360)

    st.markdown("---")

    # ── Buyer's Bank Breakdown ────────────────────────────────────────────
    buyer_bank = (df_daily.groupby("Bank Name")
                    .agg(Value=("Invoice Value","sum"), Submissions=("LC No","count"))
                    .reset_index().sort_values("Value", ascending=False).head(23))
    b1, b2 = st.columns(2)
    with b1:
        sh("🏛 Buyer's Bank Breakdown")
        fig_bb = px.bar(buyer_bank, x="Value", y="Bank Name", orientation="h",
                        color_discrete_sequence=[C[4]],
                        text=buyer_bank["Value"].map(usd))
        fig_bb.update_traces(textposition="outside", textfont_size=8, marker_color=C[4])
        fig_bb.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=500)
        st.plotly_chart(fig_bb, use_container_width=True)
    with b2:
        sh("🏛 Buyer's Bank Table")
        tbob = buyer_bank.copy()
        total_bb = tbob["Value"].sum()
        tbob["% Share"] = (tbob["Value"] / total_bb * 100).map(lambda x: f"{x:.1f}%")
        tbob["Value"]   = tbob["Value"].map(lambda x: f"${x:,.2f}")
        tbob.columns    = ["Bank Name","Invoice Value","Submissions","% Share"]
        st.dataframe(tbob, use_container_width=True, hide_index=True, height=500)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with t_overview:
    l, r = st.columns(2)
    with l:
        sh("📅 Monthly Submission Trend")
        fig = go.Figure()
        fig.add_bar(x=monthly["Month"], y=monthly["Count"], name="Submissions",
                    marker_color="#1a8fff", yaxis="y1")
        fig.add_scatter(x=monthly["Month"], y=monthly["Inv"], name="Invoice Value",
                        mode="lines+markers", line=dict(color="#00c9a7", width=2.5),
                        marker=dict(size=6), yaxis="y2")
        fig.update_layout(**PL_GENERAL,
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            yaxis2=dict(title="Invoice Value (USD)", overlaying="y", side="right",
                        gridcolor="rgba(0,0,0,0)", tickformat="$.2s"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#8899aa", size=10)),
            height=340)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("🏦 Our Bank — Invoice Value")
        fig2 = px.bar(by_bank, x="Our Bank", y="Inv", color="Our Bank",
                      color_discrete_sequence=C, text=by_bank["Inv"].apply(usd))
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(**PL, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    a, b_col = st.columns(2)
    with a:
        sh("🏢 Firm-wise Invoice Value")
        fig3 = px.pie(by_firm, names="Firm Name", values="Inv",
                      color_discrete_sequence=C, hole=0.45)
        fig3.update_layout(**PL)
        fig3.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig3, use_container_width=True)
    with b_col:
        sh("📋 Monthly Summary Table")
        tbl = monthly[["MonthSort","Month","Count","Inv","LC"]].copy()
        # `monthly` already has MonthSort in the groupby, but we rebuild `tbl` from
        # a subset. For merging breakdown values we need MonthSort back.
        # Add columns for Payment & Acceptance breakdown (based on current filters)
        # Paid            -> Payment. Rcv Dt notna
        # Accepted (Pending Pmt) -> Bank Accept Date notna AND Payment. Rcv Dt isna
        # Not Accepted    -> Bank Accept Date isna

        # Values (Invoice)
        paid_value = df[df["Payment. Rcv Dt"].notna()].groupby("MonthSort")["Invoice Value"].sum()
        pending_value = df[df["Payment. Rcv Dt"].isna()].groupby("MonthSort")["Invoice Value"].sum()
        accepted_value = df[(df["Payment. Rcv Dt"].isna()) & (df["Bank Accept Date"].notna())].groupby("MonthSort")["Invoice Value"].sum()
        not_accepted_value = df[df["Bank Accept Date"].isna()].groupby("MonthSort")["Invoice Value"].sum()

        # Counts (LC No count)
        paid_cnt = df[df["Payment. Rcv Dt"].notna()].groupby("MonthSort").size()
        accepted_cnt = df[(df["Payment. Rcv Dt"].isna()) & (df["Bank Accept Date"].notna())].groupby("MonthSort").size()
        not_accepted_cnt = df[df["Bank Accept Date"].isna()].groupby("MonthSort").size()

        tbl = tbl.merge(
            paid_value.rename("Paid Value").reset_index(),
            on="MonthSort",
            how="left",
        ).merge(
            pending_value.rename("Pending Value").reset_index(),
            on="MonthSort",
            how="left",
        ).merge(
            accepted_value.rename("Accepted Value").reset_index(),
            on="MonthSort",
            how="left",
        ).merge(
            not_accepted_value.rename("Not Accepted Value").reset_index(),
            on="MonthSort",
            how="left",
        ).merge(
            paid_cnt.rename("Paid Count").reset_index(),
            on="MonthSort",
            how="left",
        ).merge(
            accepted_cnt.rename("Accepted Count").reset_index(),
            on="MonthSort",
            how="left",
        ).merge(
            not_accepted_cnt.rename("Not Accepted Count").reset_index(),
            on="MonthSort",
            how="left",
        )

        # Fill NaNs
        for c in ["Paid Value","Pending Value","Accepted Value","Not Accepted Value","Paid Count","Accepted Count","Not Accepted Count"]:
            tbl[c] = tbl[c].fillna(0)

        # Paid Count % (requested)
        tbl["Paid Count %"] = (tbl["Paid Count"] / tbl["Count"] * 100).replace([pd.NA, pd.NaT, float("inf")], 0).round(1)

        # Rename base columns + select final display columns
        tbl = tbl[[
            "Month",
            "Count",
            "Inv",
            "LC",
            "Paid Count",
            "Paid Count %",
            "Paid Value",
            "Accepted Count",
            "Accepted Value",
            "Not Accepted Value",
        ]].copy()

        tbl.columns = [
            "Month",
            "Submissions",
            "Invoice Value (USD)",
            "LC Value (USD)",
            "Paid Count",
            "Paid Count %",
            "Paid Value (USD)",
            "Accepted Count",
            "Accepted Value (USD)",
            "Not Accepted Value (USD)",
        ]

        # Formatting
        tbl["Invoice Value (USD)"] = tbl["Invoice Value (USD)"].map(lambda x: f"${x:,.2f}")
        tbl["LC Value (USD)"]      = tbl["LC Value (USD)"].map(lambda x: f"${x:,.2f}")
        tbl["Paid Value (USD)"]    = tbl["Paid Value (USD)"].map(lambda x: f"${x:,.2f}")
        tbl["Accepted Value (USD)"] = tbl["Accepted Value (USD)"].map(lambda x: f"${x:,.2f}")
        tbl["Not Accepted Value (USD)"] = tbl["Not Accepted Value (USD)"].map(lambda x: f"${x:,.2f}")
        tbl["Paid Count"] = tbl["Paid Count"].astype(int)
        tbl["Accepted Count"] = tbl["Accepted Count"].astype(int)
        tbl["Paid Count %"] = tbl["Paid Count %"].map(lambda x: f"{x:.1f}%")

        st.dataframe(tbl, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — WEEKLY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t_weekly:
    best_wk  = weekly.loc[weekly["Count"].idxmax()]
    best_inv = weekly.loc[weekly["Inv"].idxmax()]

    w1, w2, w3, w4 = st.columns(4)
    w1.metric("📆 Total Weeks",       f"{len(weekly)}")
    w2.metric("🔝 Best Week (Subs)",  f"{best_wk['Week']} — {int(best_wk['Count'])}")
    w3.metric("💰 Best Week (Value)", f"{best_inv['Week']} — {usd(best_inv['Inv'])}")
    w4.metric("📊 Weekly Avg Subs",   f"{weekly['Count'].mean():.0f}")
    st.markdown("")

    sh("📅 Week-wise Submission Count + Invoice Value")
    fig = go.Figure()
    fig.add_bar(x=weekly["Week"], y=weekly["Count"], name="Submissions",
                marker_color="#1a8fff", yaxis="y1",
                text=weekly["Count"], textposition="outside", textfont=dict(size=9))
    fig.add_scatter(x=weekly["Week"], y=weekly["Inv"], name="Invoice Value (USD)",
                    mode="lines+markers", line=dict(color="#00c9a7", width=2.5),
                    marker=dict(size=5), yaxis="y2")
    fig.update_layout(**PL_GENERAL,
        yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
        yaxis2=dict(title="Invoice Value (USD)", overlaying="y", side="right",
                    gridcolor="rgba(0,0,0,0)", tickformat="$.2s"),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#8899aa", size=10)),
        height=340)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("---")

    l, r = st.columns(2)
    with l:
        sh("🏢 Weekly Invoice Value by Firm (Stacked)")
        fig2 = px.bar(wk_firm, x="Week", y="Inv", color="Firm Name",
                      color_discrete_sequence=C, barmode="stack")
        fig2.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig2, use_container_width=True)
    with r:
        sh("🏦 Weekly Invoice Value by Our Bank (Stacked)")
        fig3 = px.bar(wk_bank, x="Week", y="Inv", color="Our Bank",
                      color_discrete_sequence=C, barmode="stack")
        fig3.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig3, use_container_width=True)
    st.markdown("---")

    l2, r2 = st.columns(2)
    with l2:
        sh("🧑‍💼 Weekly Invoice by Sales Person (Top 6)")
        top6_sp = spg.head(6)["Sales Person"].tolist()
        wk_sp6  = wk_sp[wk_sp["Sales Person"].isin(top6_sp)]
        fig4 = px.line(wk_sp6, x="Week", y="Inv", color="Sales Person",
                       color_discrete_sequence=C, markers=True)
        fig4.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.15, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig4, use_container_width=True)
    with r2:
        sh("🔄 Weekly Payment Status (Stacked)")
        fig5 = px.bar(wk_st_grp, x="Week", y="Count", color="Status",
                      barmode="stack",
                      color_discrete_map={"Paid":"#00c9a7","Accepted":"#1a8fff","Not Accepted":"#ff6b35"})
        fig5.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig5, use_container_width=True)
    st.markdown("---")

    sh("👥 Weekly Invoice Value — Top 5 Parties (Line)")
    wk_p5 = wk_party_top[wk_party_top["Party Name"].isin(t_party.head(5)["Party Name"].tolist())]
    fig6  = px.line(wk_p5, x="Week", y="Inv", color="Party Name",
                    color_discrete_sequence=C, markers=True)
    fig6.update_layout(**PL_GENERAL,
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
        yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
        height=300)
    st.plotly_chart(fig6, use_container_width=True)
    st.markdown("---")

    sh("📋 Weekly Summary Table")

    # Weekly Summary Table: use same Status logic as “Full Record Table with Search” tab
    # Paid         -> Payment. Rcv Dt notna
    # Accepted     -> Bank Accept Date notna AND Payment. Rcv Dt isna
    # Not Accepted -> Bank Accept Date isna

    wk_status = df[["Week","Payment. Rcv Dt","Bank Accept Date","Invoice Value"]].copy()
    wk_status["Status"] = wk_status.apply(
        lambda r:
            "Paid" if pd.notna(r["Payment. Rcv Dt"]) else
            ("Accepted" if pd.notna(r["Bank Accept Date"]) else "Not Accepted"),
        axis=1
    )

    # counts
    total = wk_status.groupby("Week").size().reset_index(name="Submissions")
    paid_cnt = wk_status[wk_status["Status"] == "Paid"].groupby("Week").size().reset_index(name="Paid")

    # values by status
    val = wk_status.groupby(["Week","Status"])["Invoice Value"].sum().reset_index()
    val = val.pivot(index="Week", columns="Status", values="Invoice Value").reset_index()

    val = val.rename(columns={
        "Paid":"Paid Value",
        "Accepted":"Accepted Value",
        "Not Accepted":"Not Accepted Value",
    })

    summary = total.merge(paid_cnt, on="Week", how="left").merge(val, on="Week", how="left").fillna(0)
    summary["Payment Rate"] = (summary["Paid"] / summary["Submissions"] * 100).round(1)

    # Total invoice sum for the week (all statuses)
    total_val = wk_status.groupby("Week")["Invoice Value"].sum().reset_index(name="Invoice Value (USD)")

    summary = summary.merge(total_val, on="Week", how="left")
    summary = summary[[
        "Week",
        "Submissions",
        "Invoice Value (USD)",
        "Paid",
        "Payment Rate",
        "Paid Value",
        "Accepted Value",
        "Not Accepted Value",
    ]]

    # formatting (do this AFTER the total_val merge so symbols won't get lost)
    # ── Total sum row (requested) — compute totals BEFORE formatting ─────────────────
    # At this point, summary columns are still numeric.
    subs_total = int(summary["Submissions"].sum())
    paid_total = int(summary["Paid"].sum())

    inv_total = float(summary["Invoice Value (USD)"].sum())
    paid_val_total = float(summary["Paid Value"].sum())
    acc_val_total = float(summary["Accepted Value"].sum())
    nacc_val_total = float(summary["Not Accepted Value"].sum())

    pay_rate_total = f"{(paid_total / subs_total * 100 if subs_total else 0):.1f}%"

    totals_row = {
        "Week": "TOTAL",
        "Submissions": subs_total,
        "Invoice Value (USD)": inv_total,
        "Paid": paid_total,
        "Payment Rate": float(pay_rate_total.replace("%", "")),
        "Paid Value": paid_val_total,
        "Accepted Value": acc_val_total,
        "Not Accepted Value": nacc_val_total,
    }

    summary_total = pd.concat([summary, pd.DataFrame([totals_row])], ignore_index=True)

    # ── Formatting for display (after TOTAL row appended) ─────────────────────────────
    summary_total["Invoice Value (USD)"] = summary_total["Invoice Value (USD)"].map(lambda x: f"${x:,.2f}")
    summary_total["Paid Value"] = summary_total["Paid Value"].map(lambda x: f"${x:,.2f}")
    summary_total["Accepted Value"] = summary_total["Accepted Value"].map(lambda x: f"${x:,.2f}")
    summary_total["Not Accepted Value"] = summary_total["Not Accepted Value"].map(lambda x: f"${x:,.2f}")
    summary_total["Payment Rate"] = summary_total["Payment Rate"].map(lambda x: f"{float(x):.1f}%")
    summary_total["Paid"] = summary_total["Paid"].astype(int)

    st.dataframe(summary_total, use_container_width=True, hide_index=True)


    # ── Download Weekly Summary as PDF ────────────────────────────────
    if REPORTLAB_AVAILABLE:
        def weekly_summary_to_pdf_bytes(df_in, title="Bank Submissions Weekly Summary Table 2026",
                                          subtitle="", generated_at=""):
            buf = BytesIO()
            from reportlab.lib.pagesizes import A3, landscape as _landscape
            from reportlab.lib.units import inch as _inch

            pw, ph = _landscape(A3)
            lm = rm = 24
            tm = bm = 24
            uw = pw - lm - rm

            hf = "Helvetica-Bold"
            cf = "Helvetica"
            hfs = 12
            cfs = 9
            min_cw = 1.0 * _inch
            max_cw = 3.5 * _inch

            def mw(txt, font, sz):
                return pdfmetrics.stringWidth(str(txt), font, sz)

            # widths by content (rough)
            dt = df_in.fillna("").astype(str)
            col_widths_pdf = []
            for col in df_in.columns:
                vals = dt[col].tolist()
                if len(vals) > 200:
                    vals = vals[::max(1, len(vals)//200)]
                measured = [mw(v, cf, cfs) for v in vals if v]
                mx = max([mw(col, hf, hfs)] + measured) if measured else mw(col, hf, hfs)
                w = max(min_cw, min(max_cw, mx + 18))
                col_widths_pdf.append(w)

            tot = sum(col_widths_pdf)
            if tot > uw and tot > 0:
                col_widths_pdf = [w * uw / tot for w in col_widths_pdf]

            # split columns into groups if needed
            def chunk(cols, widths, max_w):
                groups, cur, cw3 = [], [], 0.0
                for c, w in zip(cols, widths):
                    if cur and cw3 + w > max_w:
                        groups.append(cur)
                        cur = [c]
                        cw3 = w
                    else:
                        cur.append(c)
                        cw3 += w
                if cur:
                    groups.append(cur)
                return groups

            ss2 = getSampleStyleSheet()
            hs = ParagraphStyle("HS", parent=ss2["Normal"], fontName=hf, fontSize=hfs,
                                 leading=13, textColor=colors.HexColor("#ffffff"), alignment=TA_LEFT, wordWrap="CJK")
            cs = ParagraphStyle("CS", parent=ss2["Normal"], fontName=cf, fontSize=cfs,
                                 leading=10, alignment=TA_LEFT, wordWrap="CJK")

            ts = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0d3f47")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
            ])

            col_list = list(df_in.columns)
            groups = chunk(col_list, col_widths_pdf, uw)

            pages = []
            for g_i, g_cols in enumerate(groups):
                # widths for this group
                g_w = [col_widths_pdf[col_list.index(c)] for c in g_cols]

                header_row = [Paragraph(str(c), hs) for c in g_cols]
                rows = [header_row]

                for rv in df_in.fillna("").astype(str).values.tolist():
                    row_cells = []
                    for c in g_cols:
                        idx = col_list.index(c)
                        row_cells.append(Paragraph(str(rv[idx]), cs))
                    rows.append(row_cells)

                # Center-align the table on the page. Also ensure table width is
                # aligned with available page width (uw).
                tbl2 = LongTable(
                    rows,
                    repeatRows=1,
                    colWidths=g_w,
                    hAlign="CENTER",
                    splitByRow=1,
                    spaceBefore=12,
                    spaceAfter=12,
                )

                tbl2.setStyle(ts)
                pages.append(tbl2)
                if g_i < len(groups) - 1:
                    pages.append(PageBreak())

            def page_header(canvas, doc):
                canvas.saveState()
                canvas.setFont(hf, 18)
                canvas.drawCentredString(pw/2, ph - tm - 8, str(title))

                canvas.setFont(cf, 11)
                y2 = ph - tm - 26
                if subtitle:
                    canvas.drawCentredString(pw/2, y2, str(subtitle))
                else:
                    canvas.drawString(lm, y2, "")

                if generated_at:
                    canvas.setFont(cf, 8)
                    canvas.drawString(lm, bm/2 + 2, f"Generated on: {generated_at}")

                canvas.setFont(cf, 8)
                canvas.drawRightString(pw - rm, ph - tm - 8, f"Page {doc.page}")
                canvas.restoreState()

            doc = SimpleDocTemplate(
                buf,
                pagesize=_landscape(A3),
                leftMargin=lm,
                rightMargin=rm,
                topMargin=tm + 28,
                bottomMargin=bm
            )

            doc.build(
                pages,
                onFirstPage=page_header,
                onLaterPages=page_header
            )

            buf.seek(0)
            return buf.read()

        # fixed columns (as user asked)
        # IMPORTANT: include TOTAL/sub-total row (rendered in the dataframe above)
        export_weekly = summary_total.copy()[[
            "Week",
            "Submissions",
            "Invoice Value (USD)",
            "Paid",
            "Payment Rate",
            "Paid Value",
            "Accepted Value",
            "Not Accepted Value",
        ]]

        pdf_bytes_weekly = weekly_summary_to_pdf_bytes(
            export_weekly,
            subtitle=f"Generated from current filters ({len(summary)} weeks + TOTAL row)"
                     if len(summary) else "Generated from current filters (TOTAL row only)",
            generated_at=datetime.now().strftime("%d %b %Y %H:%M:%S"),
        )

        st.download_button(
            "📄 Download Weekly Summary PDF",
            pdf_bytes_weekly,
            file_name="bank_submissions_weekly_summary.pdf",
            mime="application/pdf",
        )
    else:
        st.info("PDF download এর জন্য reportlab ইনস্টল করা নেই। requirements.txt-এ `reportlab` যোগ করুন বা `pip install reportlab` করুন।")



















# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — FIRM & SALES PERSON
# ══════════════════════════════════════════════════════════════════════════════
with t_firm:
    l, r = st.columns(2)
    with l:
        sh("🏢 Firm-wise Invoice Value")
        fig = px.bar(by_firm, y="Firm Name", x="Inv", orientation="h",
                     color="Firm Name", color_discrete_sequence=C,
                     text=by_firm["Inv"].apply(usd))
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

        sh("Firm Ranking Table")
        ft = by_firm.copy()
        ft["% Share"] = (ft["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        ft["Inv"]     = ft["Inv"].map(lambda x: f"${x:,.2f}")
        ft.columns    = ["Firm","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(ft, use_container_width=True, hide_index=True)
    with r:
        sh("🧑‍💼 Sales Person — Invoice Value (Top 12)")
        fig2 = px.bar(spg.head(12), y="Sales Person", x="Inv", orientation="h",
                      color="Sales Person", color_discrete_sequence=C,
                      text=spg.head(12)["Inv"].apply(usd))
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False, height=400)
        st.plotly_chart(fig2, use_container_width=True)

        sh("Sales Person Ranking Table")
        # Include Paid/Accepted/Not Accepted values as requested
        # Paid         -> Payment. Rcv Dt notna
        # Accepted     -> Bank Accept Date notna AND Payment. Rcv Dt isna
        # Not Accepted -> Bank Accept Date isna

        sp_rank = spg[["Sales Person","Inv","N","Paid","Pct"]].copy()

        sp_accepted = (df[df["Payment. Rcv Dt"].isna() & df["Bank Accept Date"].notna()]
                         .groupby("Sales Person")["Invoice Value"].sum())
        sp_not_accepted = (df[df["Bank Accept Date"].isna()]
                            .groupby("Sales Person")["Invoice Value"].sum())

        sp_rank["Accepted Value"] = sp_rank["Sales Person"].map(sp_accepted).fillna(0.0)
        sp_rank["Not Accepted Value"] = sp_rank["Sales Person"].map(sp_not_accepted).fillna(0.0)

        sp_paid = (df[df["Payment. Rcv Dt"].notna()]
                    .groupby("Sales Person")["Invoice Value"].sum())
        sp_rank["Paid Value"] = sp_rank["Sales Person"].map(sp_paid).fillna(0.0)

        # Format
        sp_rank["Inv"] = sp_rank["Inv"].map(lambda x: f"${x:,.2f}")
        sp_rank["Paid Value"] = sp_rank["Paid Value"].map(lambda x: f"${x:,.2f}")
        sp_rank["Accepted Value"] = sp_rank["Accepted Value"].map(lambda x: f"${x:,.2f}")
        sp_rank["Not Accepted Value"] = sp_rank["Not Accepted Value"].map(lambda x: f"${x:,.2f}")
        sp_rank["Pct"] = sp_rank["Pct"].map(lambda x: f"{x:.1f}%")
        sp_rank["Paid"] = sp_rank["Paid"].astype(int)

        st2 = sp_rank[["Sales Person","Inv","N","Paid","Pct","Paid Value","Accepted Value","Not Accepted Value"]].copy()
        st2.columns = [
            "Sales Person",
            "Invoice Value (USD)",
            "Submissions",
            "Paid",
            "Payment Rate",
            "Paid Value",
            "Accepted Value",
            "Not Accepted Value",
        ]
        st.dataframe(st2, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — BANKS
# ══════════════════════════════════════════════════════════════════════════════
with t_banks:
    l, r = st.columns(2)
    with l:
        sh("🏦 Our Bank — Share (Donut)")
        fig = px.pie(by_bank, names="Our Bank", values="Inv",
                     color_discrete_sequence=C, hole=0.48)
        fig.update_layout(**PL)
        fig.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("🏛️ Top 10 Party Banks")
        fig2 = px.bar(t_bname, y="Bank Name", x="Inv", orientation="h",
                      text=t_bname["Inv"].apply(usd), color_discrete_sequence=[C[4]])
        fig2.update_traces(textposition="outside", textfont_size=10, marker_color=C[4])
        fig2.update_layout(**PL, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    a2, b2 = st.columns(2)
    with a2:
        sh("Our Bank Detail")
        ob = by_bank.copy()
        ob["% Share"] = (ob["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        ob["Inv"]     = ob["Inv"].map(lambda x: f"${x:,.2f}")
        ob.columns    = ["Our Bank","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(ob, use_container_width=True, hide_index=True)
    with b2:
        sh("Top Party Banks Detail")
        pb = t_bname.copy()
        pb["% Share"] = (pb["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        pb["Inv"]     = pb["Inv"].map(lambda x: f"${x:,.2f}")
        pb.columns    = ["Bank Name","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(pb, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — TOP PARTIES
# ══════════════════════════════════════════════════════════════════════════════
with t_parties:
    l, r = st.columns([3, 2])
    with l:
        sh("👥 Top 10 Party — Invoice Value")
        fig = px.bar(t_party, y="Party Name", x="Inv", orientation="h",
                     color="Party Name", color_discrete_sequence=C,
                     text=t_party["Inv"].apply(usd))
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=400)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("Party Ranking")
        pt = t_party.copy()
        pt.insert(0, "#", range(1, len(pt)+1))
        pt["% Share"] = (pt["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        pt["Inv"]     = pt["Inv"].map(lambda x: f"${x:,.2f}")
        pt            = pt[["#","Party Name","Inv","N","% Share"]]
        pt.columns    = ["#","Party","Invoice Value","Subs","% Share"]
        st.dataframe(pt, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — PAYMENT STATUS
# ══════════════════════════════════════════════════════════════════════════════
with t_payment:
    s1, s2, s3 = st.columns(3)
    s1.metric("✅ Payment Received",       f"{paid_n:,}",  f"{usd(pay_v)} · {paid_n/N*100:.1f}%")
    s2.metric("⏳ Accepted — Pending Pmt", f"{acc_n:,}",   f"{usd(acc_v)} · {acc_n/N*100:.1f}%")
    s3.metric("❌ Not Accepted",           f"{nacc_n:,}",  f"{usd(nacc_v)} · {nacc_n/N*100:.1f}%")
    st.markdown("")

    st_df = pd.DataFrame({
        "Status": ["Payment Received","Accepted (Pending Pmt)","Not Accepted"],
        "Count":  [paid_n, acc_n, nacc_n],
        "Value":  [pay_v,  acc_v, nacc_v],
    })
    p1, p2 = st.columns(2)
    with p1:
        sh("By Count")
        fig = px.pie(st_df, names="Status", values="Count", hole=0.5,
                     color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"])
        fig.update_layout(**PL)
        fig.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    with p2:
        sh("By Invoice Value (USD)")
        fig2 = px.pie(st_df, names="Status", values="Value", hole=0.5,
                      color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"])
        fig2.update_layout(**PL, showlegend=False)
        fig2.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    sh("Sales Person — Paid vs Not Yet Paid")
    fig3 = go.Figure()
    fig3.add_bar(name="Paid",         x=spg["Sales Person"], y=spg["Paid"],          marker_color="#00c9a7")
    fig3.add_bar(name="Not Yet Paid", x=spg["Sales Person"], y=spg["N"]-spg["Paid"], marker_color="#095e59")
    fig3.update_layout(**PL_GENERAL, barmode="stack",
        yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
        xaxis=dict(title="", tickangle=-35, gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#8899aa", size=10)))
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    sh("Full Record Table with Search")
    search_text = st.text_input("🔍 Search all columns", key="global_search")
    pft = df.copy()
    pft["Status"] = pft.apply(lambda r:
        "✅ Paid"        if pd.notna(r["Payment. Rcv Dt"]) else
        "⏳ Accepted"   if pd.notna(r["Bank Accept Date"]) else
        "❌ Not Accepted", axis=1)
    date_cols = [
        "Bank Submition Date",
        "Bank Ref Date",
        "Lc Date",
        "Bank Accept Date",
        "Maturity Date",
        "Payment. Rcv Dt",
        "Date",
        "Invoice Date",
    ]
    date_cols = list(dict.fromkeys(date_cols))

    if search_text:
        # Match across ALL columns.
        # Only date columns are converted to the same visible format (DD MMM YYYY)
        # before matching, so user-typed dates will be found correctly.
        pft_for_search = pft.copy()

        for col in date_cols:
            if col in pft_for_search.columns:
                pft_for_search[col] = (
                    pd.to_datetime(pft_for_search[col], errors="coerce")
                    .dt.strftime("%d %b %Y")
                    .fillna("")
                )

        # Match across ALL columns as strings.
        # This ensures values like Bank Refno (even if Excel imported as numeric)
        # still match against the typed text.
        tmp = pft_for_search.copy()
        tmp = tmp.astype(str)

        # Ensure date columns use visible DD MMM YYYY representation only for matching.
        for col in date_cols:
            if col in tmp.columns:
                tmp[col] = (
                    pd.to_datetime(pft_for_search[col], errors="coerce")
                    .dt.strftime("%d %b %Y")
                    .fillna("")
                    .astype(str)
                )

        mask = tmp.apply(lambda c: c.str.contains(search_text, case=False, na=False)).any(axis=1)
        pft = pft[mask]



    # Keep datetime dtype in the visible columns so Streamlit sorting works.
    # UI format is handled by st.column_config.DateColumn.
    for col in date_cols:
        if col in pft.columns:
            dtcol = pd.to_datetime(pft[col], errors="coerce")
            pft[col + "__dt"] = dtcol
            pft[col] = dtcol



    internal_cols = ["_date","MonthSort","Month","WeekSort","Week","DayName"]
    # hidden datetime helpers for correct sorting
    hidden_dt_cols = [c + "__dt" for c in date_cols if c in pft.columns]
    internal_cols = internal_cols + hidden_dt_cols
    col_order = [
        "Firm Name","Our Bank","Bank Submition Date","Bank Ref Date","Bank Refno",
        "Party Name","LC No","Lc Date","Tenor","Bank Name","Invoice No","Invoice Date",
        "Invoice Qty","Invoice Value","Bank Accept Date","Maturity Date",
        "Payment. Rcv Dt","Sales Person","Week","DayName","Date","Status",
    ]
    export_cols = [c for c in col_order if c in pft.columns]
    extra_cols  = [c for c in pft.columns if c not in export_cols and c not in internal_cols and c not in hidden_dt_cols]
    pft_export  = pft[export_cols + extra_cols]
    pft_display = pft_export.copy()

    # Total invoice value
    total_invoice_value = 0.0
    if "Invoice Value" in pft_export.columns:
        try: total_invoice_value = float(pft_export["Invoice Value"].sum())
        except Exception: pass

    # Keep numeric values for proper sorting.
    # Display formatting will be handled via st.column_config below.
    # (Do NOT convert Invoice Value to currency string here.)

    # Column widths
    def make_col_widths(dataframe, min_w=100, max_w=450, cw=8):
        text_df = dataframe.fillna("").astype(str)
        widths  = {}
        for col in text_df.columns:
            try:
                max_len = float(text_df[col].str.len().max())
                if pd.isna(max_len):
                    max_len = 0.0
            except Exception:
                max_len = 0.0
            w = max(len(str(col)), int(max_len)) * cw + 24
            widths[col] = min(max_w, max(min_w, w))
        if "Bank Refno"    in widths: widths["Bank Refno"]    = max(widths["Bank Refno"],    285)
        if "LC No"    in widths: widths["LC No"]    = max(widths["LC No"],    285)
        if "Invoice Value" in widths: widths["Invoice Value"] = max(widths["Invoice Value"], 160)
        return widths

    col_widths  = make_col_widths(pft_display)

    # Use numeric columns for proper ascending/descending sorting.
    # For currency-like columns, keep numbers and only apply formatting via column_config.
    col_config = {}
    date_cols_set = set(date_cols)
    for c in pft_display.columns:
        if c == "Invoice Value" and pd.api.types.is_numeric_dtype(pft_display[c]):
            col_config[c] = st.column_config.NumberColumn(
                format="$%0.2f",
                width=col_widths[c],
            )
        elif c in date_cols_set:
            # Sort uses datetime dtype; format for display.
            col_config[c] = st.column_config.DateColumn(
                width=col_widths[c],
                format="DD MMM YYYY",
            )
        elif pd.api.types.is_numeric_dtype(pft_display[c]):
            col_config[c] = st.column_config.NumberColumn(width=col_widths[c])
        else:
            col_config[c] = st.column_config.TextColumn(width=col_widths[c])


    st.dataframe(pft_display, use_container_width=True, hide_index=True,
                 height=500, column_config=col_config)
    if "Invoice Value" in pft_export.columns:
        st.markdown(f"**Total Invoice Value (filtered):** ${total_invoice_value:,.2f}")

    # ── Downloads ──────────────────────────────────────────────────────────
    col_csv, col_pdf = st.columns(2)
    col_csv.download_button("📥 Download CSV",
        pft_export.to_csv(index=False).encode("utf-8"),
        "bank_submit_filtered.csv", "text/csv")

    pdf_width_mode = st.selectbox("PDF column width mode",
        ["Auto (by content)", "Equal", "Custom (inches, comma-separated)"])
    custom_widths_input = ""
    if pdf_width_mode == "Custom (inches, comma-separated)":
        custom_widths_input = st.text_input("Custom widths (e.g. 1.0,2.5,1.5)")

    pdf_period = (f"{date_range[0].strftime('%d %b %Y')} – {date_range[1].strftime('%d %b %Y')}"
                  if isinstance(date_range, tuple) and len(date_range) == 2 else str(date_range))

    if REPORTLAB_AVAILABLE:
        def df_to_pdf_bytes(df_in, title="Bank submit status", subtitle="",
                            custom_widths=None, generated_at=""):
            buf = BytesIO()
            lm = rm = tm = bm = 24
            pw, ph = landscape(A3)
            uw = pw - lm - rm
            hf = "Helvetica-Bold"; cf = "Helvetica"; hfs = 10; cfs = 8
            min_cw = 1.2*inch; max_cw = 4.5*inch

            def mw(txt, font, sz): return pdfmetrics.stringWidth(str(txt), font, sz)

            col_widths_pdf = []
            if custom_widths and isinstance(custom_widths, (list, tuple)):
                col_widths_pdf = [max(min_cw, min(max_cw, float(w)*inch)) for w in custom_widths]
                if len(col_widths_pdf) < len(df_in.columns):
                    rem = len(df_in.columns) - len(col_widths_pdf)
                    add = max(0, uw - sum(col_widths_pdf)) / rem if rem else min_cw
                    col_widths_pdf.extend([max(min_cw, min(max_cw, add))]*rem)
                tot = sum(col_widths_pdf)
                if tot > uw and tot > 0:
                    col_widths_pdf = [w*uw/tot for w in col_widths_pdf]
            elif custom_widths == "EQUAL":
                col_widths_pdf = [uw/len(df_in.columns)]*len(df_in.columns)
            else:
                dt = df_in.fillna("").astype(str)
                for col in df_in.columns:
                    hw = mw(col, hf, hfs)
                    vals = dt[col].tolist()
                    if len(vals) > 250:
                        vals = vals[::max(1, len(vals)//250)]
                    measured = sorted([mw(v, cf, cfs) for v in vals if v])
                    cw2 = measured[min(len(measured)-1, int(len(measured)*0.9))] if measured else mw("M", cf, cfs)
                    col_widths_pdf.append(min(max_cw, max(min_cw, max(hw, cw2)+16)))
                tot = sum(col_widths_pdf)
                if tot > uw and tot > 0:
                    col_widths_pdf = [w*uw/tot for w in col_widths_pdf]
                elif tot < uw and tot > 0:
                    col_widths_pdf = [w + (uw-tot)*(w/tot) for w in col_widths_pdf]

            ss2 = getSampleStyleSheet()
            hs  = ParagraphStyle("H", parent=ss2["Normal"], fontName=hf, fontSize=hfs,
                                 leading=11, textColor=colors.white, alignment=TA_LEFT, wordWrap="CJK")
            cs  = ParagraphStyle("C", parent=ss2["Normal"], fontName=cf, fontSize=cfs,
                                 leading=10, alignment=TA_LEFT, wordWrap="CJK")

            def chunk(cols, widths, max_w):
                groups, cur, cw3 = [], [], 0.0
                for c, w in zip(cols, widths):
                    if cur and cw3 + w > max_w: groups.append(cur); cur=[c]; cw3=w
                    else: cur.append(c); cw3+=w
                if cur: groups.append(cur)
                return groups

            ts = TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0d3f47")),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("ALIGN",(0,0),(-1,-1),"LEFT"),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
                ("FONTSIZE",(0,0),(-1,-1),8),
                ("LEFTPADDING",(0,0),(-1,-1),5),
                ("RIGHTPADDING",(0,0),(-1,-1),5),
                ("BOTTOMPADDING",(0,0),(-1,-1),4),
                ("TOPPADDING",(0,0),(-1,-1),4),
                ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                ("BOX",(0,0),(-1,-1),0.5,colors.black),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.whitesmoke,colors.lightgrey]),
            ])

            pages = []
            for g_i, g_cols in enumerate(chunk(list(df_in.columns), col_widths_pdf, uw)):
                g_w = [col_widths_pdf[list(df_in.columns).index(c)] for c in g_cols]
                rows = [[Paragraph(str(c), hs) for c in g_cols]]
                for rv in df_in.fillna("").astype(str).values.tolist():
                    rows.append([Paragraph(str(rv[list(df_in.columns).index(c)]), cs) for c in g_cols])
                tbl2 = LongTable(rows, repeatRows=1, colWidths=g_w, hAlign="LEFT",
                                 splitByRow=1, spaceBefore=12, spaceAfter=12)
                tbl2.setStyle(ts)
                pages.append(tbl2)
                if g_i < len(chunk(list(df_in.columns), col_widths_pdf, uw))-1:
                    pages.append(PageBreak())

            def ph2(canvas, doc):
                canvas.saveState()
                ty=ph-tm+10; sy=ty-16; gy=sy-14
                canvas.setFont(hf,20); canvas.drawCentredString(pw/2,ty,str(title))
                canvas.setFont(cf,11); canvas.drawCentredString(pw/2,sy,str(subtitle))
                if generated_at:
                    canvas.setFont(cf,8); canvas.drawString(lm,gy,f"Generated on: {generated_at}")
                canvas.setFont(cf,8); canvas.drawRightString(pw-rm,ty,f"Page {doc.page}")
                canvas.setFont(hf,8); canvas.drawRightString(pw-rm,bm/2,"ASM")
                canvas.restoreState()

            doc = SimpleDocTemplate(buf, pagesize=landscape(A3),
                leftMargin=lm, rightMargin=rm, topMargin=tm+28, bottomMargin=bm)
            doc.build(pages, onFirstPage=ph2, onLaterPages=ph2)
            buf.seek(0); return buf.read()

        try:
            parsed_custom = None
            if pdf_width_mode == "Auto (by content)":
                tw = [max(1.2, min(4.5, col_widths.get(c, 100)/96.0)) for c in pft_display.columns]
                parsed_custom = tw
            elif pdf_width_mode == "Equal":
                parsed_custom = "EQUAL"
            elif pdf_width_mode == "Custom (inches, comma-separated)" and custom_widths_input:
                try: parsed_custom = [float(x.strip()) for x in custom_widths_input.split(",") if x.strip()]
                except Exception: col_pdf.error("Invalid custom widths.")

            pdf_df = pft_export.copy()
            if "Invoice Value" in pdf_df.columns:
                pdf_df["Invoice Value"] = pdf_df["Invoice Value"].map(lambda x: f"${x:,.2f}")
            if len(pdf_df.columns) > 0:
                tr = {c: "" for c in pdf_df.columns}
                tr[pdf_df.columns[0]] = "TOTAL"
                if "Invoice Value" in pdf_df.columns:
                    tr["Invoice Value"] = f"${total_invoice_value:,.2f}"
                pdf_df = pd.concat([pdf_df, pd.DataFrame([tr])], ignore_index=True)

            pdf_bytes = df_to_pdf_bytes(pdf_df, subtitle=pdf_period,
                custom_widths=parsed_custom, generated_at=datetime.now().strftime("%d %b %Y %H:%M:%S"))
            col_pdf.download_button("📄 Download PDF", pdf_bytes,
                "bank_submit_status.pdf", "application/pdf")
        except Exception as e:
            col_pdf.error(f"Could not generate PDF: {e}")
    else:
        col_pdf.warning("Install reportlab: pip install reportlab")

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#445566;font-size:11px;letter-spacing:.1em;'>"
    f"Asm@2026  BANK SUBMIT HISTORY DASHBOARD &nbsp;·&nbsp; {N:,} RECORDS &nbsp;·&nbsp; {period}"
    f"</p>", unsafe_allow_html=True)
