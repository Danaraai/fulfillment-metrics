"""
dashboard.py  —  FC1 Performance Dashboard
Run locally:  streamlit run dashboard.py
Deploy:       push to GitHub → connect on share.streamlit.io
"""

import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, date

from data_loader import load_export, load_labor_hours, load_comparison, _comparison_load_error

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="FC1 Metrics",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────

st.markdown("""
<style>
  /* ── Global background ── */
  .stApp { background-color: #f8f9fb; }

  /* ── Sidebar ── */
  [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e5e9f0; }

  /* ── KPI cards ── */
  [data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #e5e9f0;
    border-radius: 12px;
    padding: 20px 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  }
  [data-testid="metric-container"] label {
    color: #6b7a99 !important;
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  [data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 30px; font-weight: 700; color: #1a202c;
  }

  /* ── Section headers ── */
  .section-header {
    font-size: 15px; font-weight: 700; color: #2d3748;
    margin: 28px 0 10px; padding-bottom: 8px;
    border-bottom: 2px solid #e5e9f0;
    text-transform: uppercase; letter-spacing: 0.04em;
  }

  /* ── Dataframe ── */
  [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

  /* ── Title ── */
  h1 { color: #1a202c !important; font-weight: 800 !important; }
</style>
""", unsafe_allow_html=True)

PLOTLY_THEME = dict(
    plot_bgcolor="#ffffff",
    paper_bgcolor="#f8f9fb",
    font_color="#2d3748",
    xaxis=dict(gridcolor="#e5e9f0", linecolor="#e5e9f0", showgrid=True),
    yaxis=dict(gridcolor="#e5e9f0", linecolor="#e5e9f0", showgrid=True),
)

# Carrier colors match the transit time chart (Amazon=orange, FedEx=purple, UPS=blue, USPS=green)
CARRIER_COLORS = {
    "Amazon Shipping": "#f8961e",
    "FedEx":           "#b794f4",
    "UPS":             "#4361ee",
    "USPS":            "#90c97a",
}
FALLBACK_COLORS = ["#4361ee", "#f72585", "#4cc9f0", "#f8961e", "#7209b7", "#3a86ff"]

LABOR_RATE   = 18.25   # $/hr
PKG_COST     = 0.31    # $/order before 2026-03-19
PKG_COST_NEW = 0.28    # $/order from 2026-03-19 onwards
PKG_CUTOVER  = pd.Timestamp("2026-03-19")

# ── Load data ─────────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def get_data():
    return load_export(), load_labor_hours()

with st.spinner("Loading data from Google Sheets…"):
    export_df, labor_df = get_data()

with st.spinner("Loading negotiation comparison data…"):
    comparison_df = load_comparison()

if export_df.empty:
    st.error("No shipping data found in the Export tab. Please run the weekly export first.")
    st.stop()

# ── Sidebar filters ───────────────────────────────────────────────────────────

st.sidebar.title("📦 FC1 Metrics")
st.sidebar.markdown("---")

min_date = export_df["Transaction Date"].min().date()
max_date = export_df["Transaction Date"].max().date()

default_start = max(min_date, datetime.now().date().replace(month=1, day=1))

date_range = st.sidebar.date_input(
    "Date range",
    value=(default_start, max_date),
    min_value=date(2024, 1, 1),
    max_value=max_date,
)

if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = default_start, max_date

granularity = st.sidebar.selectbox("Chart granularity", ["Weekly", "Monthly"], index=0)

# Apply date filter to export
mask = (
    (export_df["Transaction Date"].dt.date >= start_date) &
    (export_df["Transaction Date"].dt.date <= end_date)
)
df = export_df[mask].copy()

# Per-order packaging cost (rate changed 2026-03-19)
df["Pkg Cost"] = df["Transaction Date"].apply(
    lambda d: PKG_COST_NEW if d >= PKG_CUTOVER else PKG_COST
)

# True shipping cost = sum of all individual cost components
# (Original Invoice omits WMS Fuel Surcharge, Delivery Area Surcharge, Insurance, etc.)
_SHIP_COLS = ["Fulfillment without Surcharge", "Surcharge Applied",
              "WMS Fuel Surcharge", "Delivery Area Surcharge",
              "Residential Area Surcharge", "Address Correction Fee",
              "Other Order Fee", "Insurance Amount"]
df["Total Ship Cost"] = sum(df[c] for c in _SHIP_COLS if c in df.columns)

# Apply date filter to labor
if not labor_df.empty:
    l_mask = (
        (labor_df["Date"].dt.date >= start_date) &
        (labor_df["Date"].dt.date <= end_date)
    )
    ldf = labor_df[l_mask].copy()
else:
    ldf = pd.DataFrame()

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing {len(df):,} shipments  \n{start_date} → {end_date}")
st.sidebar.caption(f"Last refreshed: {datetime.now().strftime('%b %d, %Y %H:%M')}")

if st.sidebar.button("🔄 Refresh data"):
    st.cache_data.clear()
    st.rerun()

# ── Period helper ─────────────────────────────────────────────────────────────
# Returns the period-start timestamp for each row based on granularity setting.

def week_start(series: pd.Series) -> pd.Series:
    """Monday-anchored week start, normalized (no time component)."""
    return (series - pd.to_timedelta(series.dt.dayofweek, unit="D")).dt.normalize()

def period_start(series: pd.Series) -> pd.Series:
    """Return week or month start depending on sidebar granularity selection."""
    if granularity == "Monthly":
        return series.dt.to_period("M").dt.to_timestamp()
    return week_start(series)

def period_label(ts_series: pd.Series) -> pd.Series:
    """Format period-start timestamps as display labels."""
    if granularity == "Monthly":
        return ts_series.dt.strftime("%b '%y")   # e.g. "Oct '25"
    return ts_series.dt.strftime("%-m-%d")        # e.g. "10-27"

df["Week"] = period_start(df["Transaction Date"])

# ── Header ────────────────────────────────────────────────────────────────────

st.title("📦 FC1 Performance Dashboard")
st.caption(f"Jack Archer merchant  ·  {start_date.strftime('%b %d, %Y')} – {end_date.strftime('%b %d, %Y')}")

# ── Build daily table (used for KPIs + daily table section) ──────────────────

# Daily order counts
day_counts = (
    df["Transaction Date"].dt.date
      .value_counts()
      .sort_index()
      .reset_index()
)
day_counts.columns = ["Date", "Daily Orders"]
day_counts["Date"] = pd.to_datetime(day_counts["Date"])

# Daily avg shipping cost (sum of all surcharge components)
if "Total Ship Cost" in df.columns:
    avg_cost_daily = (
        df.assign(_day=df["Transaction Date"].dt.date)
          .groupby("_day")["Total Ship Cost"]
          .mean()
          .reset_index()
          .rename(columns={"_day": "Date", "Total Ship Cost": "Avg_Ship_Cost"})
    )
    avg_cost_daily["Date"] = pd.to_datetime(avg_cost_daily["Date"])
    daily_orders = day_counts.merge(avg_cost_daily, on="Date", how="left")
else:
    daily_orders = day_counts.copy()
    daily_orders["Avg_Ship_Cost"] = float("nan")


# Join labor hours (deduplicated)
if not ldf.empty:
    labor_clean = (
        ldf[["Date", "Outbound Hours", "Total Hours", "Emp Hours", "Temp Hours", "Headcount"]]
        .copy()
        .assign(Date=pd.to_datetime(ldf["Date"].dt.date))
        .drop_duplicates(subset=["Date"], keep="last")
    )
    daily_tbl = daily_orders.merge(labor_clean, on="Date", how="left")
else:
    daily_tbl = daily_orders.copy()
    for col in ["Outbound Hours", "Total Hours", "Emp Hours", "Temp Hours", "Headcount"]:
        daily_tbl[col] = float("nan")

# Compute daily metrics
daily_tbl["OPLH"] = (
    daily_tbl["Daily Orders"] / daily_tbl["Total Hours"]
).where(daily_tbl["Total Hours"].fillna(0) > 0)

daily_tbl["Total Labor Cost/Order"] = (
    daily_tbl["Total Hours"] * LABOR_RATE / daily_tbl["Daily Orders"]
).where(daily_tbl["Daily Orders"] > 0)

daily_tbl["Outbound Labor Cost/Order"] = (
    daily_tbl["Outbound Hours"] * LABOR_RATE / daily_tbl["Daily Orders"]
).where(daily_tbl["Daily Orders"] > 0)

daily_tbl["Avg Shipping Cost/Order"] = daily_tbl["Avg_Ship_Cost"]

# Daily packaging cost — rate changed 2026-03-19
pkg_daily = (
    df.assign(_day=df["Transaction Date"].dt.date)
      .groupby("_day")["Pkg Cost"]
      .mean()
      .reset_index()
      .rename(columns={"_day": "Date", "Pkg Cost": "Pkg Cost/Order"})
)
pkg_daily["Date"] = pd.to_datetime(pkg_daily["Date"])
daily_tbl = daily_tbl.merge(pkg_daily, on="Date", how="left")
daily_tbl["Pkg Cost/Order"] = daily_tbl["Pkg Cost/Order"].fillna(PKG_COST)

daily_tbl["Total Cost/Order"] = (
    daily_tbl["Total Labor Cost/Order"].fillna(0)
    + daily_tbl["Pkg Cost/Order"]
    + daily_tbl["Avg Shipping Cost/Order"].fillna(0)
)

# ── Build WEEKLY aggregation ──────────────────────────────────────────────────

# Weekly orders + shipping from export
weekly_export = (
    df.groupby("Week")
      .agg(
          Orders=("Transaction Date", "count"),
          Avg_Ship=("Total Ship Cost", "mean"),
          Avg_Pkg=("Pkg Cost", "mean"),
      )
      .reset_index()
)

# Weekly labor from ldf
if not ldf.empty:
    ldf2 = ldf.copy()
    ldf2["Week"] = period_start(ldf2["Date"])
    weekly_labor = (
        ldf2.groupby("Week")
            .agg(
                Total_Hours=("Total Hours", "sum"),
                Outbound_Hours=("Outbound Hours", "sum"),
            )
            .reset_index()
    )
    weekly = weekly_export.merge(weekly_labor, on="Week", how="left")
else:
    weekly = weekly_export.copy()
    weekly["Total_Hours"]    = float("nan")
    weekly["Outbound_Hours"] = float("nan")

weekly["OPLH"] = (weekly["Orders"] / weekly["Total_Hours"]).where(weekly["Total_Hours"].fillna(0) > 0)
weekly["Total Labor Cost/Order"]    = (weekly["Total_Hours"]    * LABOR_RATE / weekly["Orders"]).where(weekly["Orders"] > 0)
weekly["Outbound Labor Cost/Order"] = (weekly["Outbound_Hours"] * LABOR_RATE / weekly["Orders"]).where(weekly["Orders"] > 0)
weekly["Pkg Cost/Order"]            = weekly["Avg_Pkg"].fillna(PKG_COST)
weekly["Total Cost/Order"]          = (
    weekly["Total Labor Cost/Order"].fillna(0)
    + weekly["Pkg Cost/Order"]
    + weekly["Avg_Ship"].fillna(0)
)
weekly["Week Label"] = period_label(weekly["Week"])

# ── KPI cards ─────────────────────────────────────────────────────────────────

total_orders  = len(df)
avg_ship_cost = df["Total Ship Cost"].mean() if "Total Ship Cost" in df.columns else None
_has_labor     = daily_tbl["Total Hours"].fillna(0) > 0
_total_hrs     = daily_tbl.loc[_has_labor, "Total Hours"].sum()
_total_ord_lbr = daily_tbl.loc[_has_labor, "Daily Orders"].sum()
avg_oplh       = (_total_ord_lbr / _total_hrs)              if _total_hrs > 0     else float("nan")
avg_labor_cost = (_total_hrs * LABOR_RATE / _total_ord_lbr) if _total_ord_lbr > 0 else float("nan")

avg_pkg_cost   = df["Pkg Cost"].mean() if "Pkg Cost" in df.columns else PKG_COST
avg_total_cost = (
    (avg_labor_cost or 0) + avg_pkg_cost + (avg_ship_cost or 0)
) if pd.notna(avg_labor_cost) and pd.notna(avg_ship_cost) else None

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Shipments", f"{total_orders:,}")
k2.metric("Avg Shipping Cost / Order",
          f"${avg_ship_cost:.2f}" if pd.notna(avg_ship_cost) and avg_ship_cost != 0 else "—")
k3.metric("Avg OPLH (Total Hours)",
          f"{avg_oplh:.1f}" if pd.notna(avg_oplh) else "—",
          help="Orders Per Total Labor Hour")
k4.metric("Avg Labor Cost / Order",
          f"${avg_labor_cost:.2f}" if pd.notna(avg_labor_cost) else "—")
k5.metric("Avg Total Cost / Order",
          f"${avg_total_cost:.2f}" if avg_total_cost is not None else "—",
          help="Labor + Packaging + Shipping")

# ── Chart 1: Weekly shipments bar ─────────────────────────────────────────────

st.markdown('<div class="section-header">Total # of Shipments Fulfilled by FC1</div>',
            unsafe_allow_html=True)

fig_bar_vol = go.Figure(go.Bar(
    x=weekly["Week Label"],
    y=weekly["Orders"],
    text=weekly["Orders"],
    textposition="outside",
    textfont=dict(size=12, color="#4361ee", family="monospace"),
    marker_color="#4361ee",
    marker_line_width=0,
))
fig_bar_vol.update_layout(
    **PLOTLY_THEME,
    xaxis_title="Week",
    yaxis_title="",
    margin=dict(t=30, b=10, l=10, r=10),
    height=380,
    uniformtext_minsize=10,
    uniformtext_mode="hide",
)
fig_bar_vol.update_xaxes(type="category")   # prevent Plotly auto-parsing M-DD as dates
st.plotly_chart(fig_bar_vol, use_container_width=True)

# ── Chart 2: Weekly cost summary TABLE ───────────────────────────────────────

st.markdown('<div class="section-header">Weekly Cost Summary</div>', unsafe_allow_html=True)

tbl = weekly[["Week Label", "OPLH", "Total Labor Cost/Order",
              "Pkg Cost/Order", "Outbound Labor Cost/Order",
              "Avg_Ship", "Total Cost/Order"]].copy()
tbl.columns = [
    "Week",
    "Avg OPLH",
    "Total Labor Cost/Order",
    "Packaging Cost/Order",
    "Outbound Labor Cost/Order",
    "Avg Shipping Cost w/ Surcharges",
    "Cost per Order (Pick, Pack, Ship)",
]

# Grand total row — use weighted totals, not unweighted mean of weekly ratios
_w_orders   = weekly["Orders"].sum()
_w_tot_hrs  = weekly["Total_Hours"].fillna(0).sum()
_w_out_hrs  = weekly["Outbound_Hours"].fillna(0).sum()
_w_avg_ship = tbl["Avg Shipping Cost w/ Surcharges"].mean()   # avg of weekly avgs is appropriate here
_w_avg_pkg  = df["Pkg Cost"].mean() if "Pkg Cost" in df.columns else PKG_COST
_w_tot_lbr  = (_w_tot_hrs * LABOR_RATE / _w_orders) if _w_orders > 0 else float("nan")
_w_out_lbr  = (_w_out_hrs * LABOR_RATE / _w_orders) if _w_orders > 0 else float("nan")
_w_oplh     = (_w_orders / _w_tot_hrs)              if _w_tot_hrs > 0 else float("nan")
grand = {
    "Week": "Grand Total",
    "Avg OPLH":                          _w_oplh,
    "Total Labor Cost/Order":            _w_tot_lbr,
    "Packaging Cost/Order":              _w_avg_pkg,
    "Outbound Labor Cost/Order":         _w_out_lbr,
    "Avg Shipping Cost w/ Surcharges":   _w_avg_ship,
    "Cost per Order (Pick, Pack, Ship)": ((_w_tot_lbr or 0) + _w_avg_pkg + (_w_avg_ship or 0)),
}
tbl_display = pd.concat([tbl, pd.DataFrame([grand])], ignore_index=True)

money_cols = [
    "Total Labor Cost/Order",
    "Packaging Cost/Order",
    "Outbound Labor Cost/Order",
    "Avg Shipping Cost w/ Surcharges",
    "Cost per Order (Pick, Pack, Ship)",
]
fmt_tbl = {"Avg OPLH": "{:.0f}"}
for c in money_cols:
    fmt_tbl[c] = "${:.2f}"

def _style_grand(row):
    if row["Week"] == "Grand Total":
        return ["font-weight: bold; background-color: #f0f4ff"] * len(row)
    return [""] * len(row)

st.dataframe(
    tbl_display.style
        .format(fmt_tbl, na_rep="—")
        .apply(_style_grand, axis=1),
    use_container_width=True,
    hide_index=True,
    height=min(450, (len(tbl_display) + 1) * 38 + 40),
)

# ── Chart 3: Cost per order stacked bar ──────────────────────────────────────

st.markdown('<div class="section-header">Cost per Order Breakdown</div>', unsafe_allow_html=True)
st.caption("⚠️ Surcharges included in shipping cost (Original Invoice)")

cost_chart = weekly[weekly["Total Cost/Order"].notna()].copy()

fig_stack = go.Figure()

# Bottom: Total Labor Cost (teal)
fig_stack.add_bar(
    x=cost_chart["Week Label"],
    y=cost_chart["Total Labor Cost/Order"].round(2),
    name="Total Labor Cost per Order",
    marker_color="#4cc9f0",
    text=cost_chart["Total Labor Cost/Order"].round(2),
    texttemplate="%{text}",
    textposition="inside",
    insidetextanchor="middle",
    textfont=dict(size=11, color="#1a202c"),
)

# Middle: Packaging (yellow)
fig_stack.add_bar(
    x=cost_chart["Week Label"],
    y=[PKG_COST] * len(cost_chart),
    name="Packaging Cost per Order",
    marker_color="#ffd166",
    text=[PKG_COST] * len(cost_chart),
    texttemplate="%{text:.2f}",
    textposition="inside",
    insidetextanchor="middle",
    textfont=dict(size=11, color="#1a202c"),
)

# Top: Avg Shipping (green)
fig_stack.add_bar(
    x=cost_chart["Week Label"],
    y=cost_chart["Avg_Ship"].round(2),
    name="Avg Shipping Cost w/ Surcharges",
    marker_color="#8db89c",
    text=cost_chart["Avg_Ship"].round(2),
    texttemplate="%{text:.2f}",
    textposition="inside",
    insidetextanchor="middle",
    textfont=dict(size=11, color="#ffffff"),
)

# Total labels above each bar
totals = (
    cost_chart["Total Labor Cost/Order"].fillna(0)
    + PKG_COST
    + cost_chart["Avg_Ship"].fillna(0)
).round(2)

fig_stack.add_scatter(
    x=cost_chart["Week Label"],
    y=totals + 0.3,
    mode="text",
    text=[f"${v:.2f}" for v in totals],
    textfont=dict(size=12, color="#1a202c", family="monospace"),
    showlegend=False,
)

fig_stack.update_layout(
    **PLOTLY_THEME,
    barmode="stack",
    xaxis_title="Week",
    yaxis_title="Cost per Order ($)",
    yaxis_tickprefix="$",
    legend=dict(orientation="h", yanchor="top", y=-0.18, x=0.5, xanchor="center"),
    margin=dict(t=30, b=80, l=10, r=10),
    height=420,
)
fig_stack.update_xaxes(type="category")   # prevent Plotly auto-parsing M-DD as dates
st.plotly_chart(fig_stack, use_container_width=True)

# ── Chart 4: Avg Transit Time by Carrier ─────────────────────────────────────

if "Transit Time (Days)" in df.columns and "Carrier" in df.columns:
    st.markdown('<div class="section-header">Avg Transit Time per Carrier (Days)</div>',
                unsafe_allow_html=True)

    transit = (
        df[df["Transit Time (Days)"] > 0]
          .groupby(["Week", "Carrier"])["Transit Time (Days)"]
          .mean()
          .reset_index()
    )
    transit["Week Label"] = period_label(transit["Week"])
    transit = transit.sort_values("Week")

    carriers = sorted(transit["Carrier"].unique())
    color_seq = [CARRIER_COLORS.get(c, FALLBACK_COLORS[i % len(FALLBACK_COLORS)])
                 for i, c in enumerate(carriers)]

    fig_transit = px.line(
        transit,
        x="Week Label",
        y="Transit Time (Days)",
        color="Carrier",
        color_discrete_sequence=color_seq,
        markers=True,
        category_orders={"Carrier": carriers},
    )
    fig_transit.update_traces(line_width=2.5, marker_size=6)
    fig_transit.update_layout(
        **PLOTLY_THEME,
        xaxis_title="Week",
        yaxis_title="Avg Transit Days",
        legend=dict(orientation="h", yanchor="top", y=-0.18, x=0.5, xanchor="center"),
        margin=dict(t=20, b=80, l=10, r=10),
        height=400,
    )
    fig_transit.update_xaxes(type="category")   # prevent "11-03" being parsed as year 2011
    st.plotly_chart(fig_transit, use_container_width=True)

# ── Chart 5: Carrier mix + avg cost ──────────────────────────────────────────

st.markdown('<div class="section-header">Carrier Performance</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)

if "Carrier" in df.columns:
    carrier_counts = (
        df.groupby("Carrier")
          .size()
          .reset_index(name="Shipments")
          .sort_values("Shipments", ascending=False)
    )
    pie_colors = [CARRIER_COLORS.get(c, FALLBACK_COLORS[i % len(FALLBACK_COLORS)])
                  for i, c in enumerate(carrier_counts["Carrier"])]

    with c1:
        st.markdown("**Carrier Mix**")
        fig_pie = px.pie(
            carrier_counts,
            names="Carrier",
            values="Shipments",
            color_discrete_sequence=pie_colors,
            hole=0.4,
        )
        fig_pie.update_layout(**PLOTLY_THEME, showlegend=True,
                               margin=dict(t=10, b=10, l=10, r=10))
        fig_pie.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        st.markdown("**Avg Shipping Cost by Carrier**")
        carrier_cost = (
            df.groupby("Carrier")["Original Invoice"]
              .mean()
              .reset_index()
              .rename(columns={"Original Invoice": "Avg Cost ($)"})
              .sort_values("Avg Cost ($)", ascending=True)
        )
        bar_colors = [CARRIER_COLORS.get(c, FALLBACK_COLORS[i % len(FALLBACK_COLORS)])
                      for i, c in enumerate(carrier_cost["Carrier"])]
        fig_hbar = go.Figure(go.Bar(
            y=carrier_cost["Carrier"],
            x=carrier_cost["Avg Cost ($)"],
            orientation="h",
            text=carrier_cost["Avg Cost ($)"].map("${:.2f}".format),
            textposition="outside",
            marker_color=bar_colors,
        ))
        fig_hbar.update_layout(
            **PLOTLY_THEME,
            xaxis_tickprefix="$",
            margin=dict(t=10, b=10, l=10, r=10),
            height=300,
        )
        st.plotly_chart(fig_hbar, use_container_width=True)

# ── Daily Operations Metrics Table ────────────────────────────────────────────

st.markdown('<div class="section-header">Daily Operations Metrics</div>', unsafe_allow_html=True)

show = daily_tbl.rename(columns={
    "Outbound Hours": "Labor Hrs Outbound",
    "Total Hours":    "Labor Hrs Total",
    "Emp Hours":      "Emp Hrs",
    "Temp Hours":     "Temp Hrs",
})[[
    "Date", "Daily Orders",
    "Labor Hrs Outbound", "Labor Hrs Total", "Emp Hrs", "Temp Hrs", "Headcount",
    "OPLH",
    "Total Labor Cost/Order", "Outbound Labor Cost/Order",
    "Avg Shipping Cost/Order", "Pkg Cost/Order", "Total Cost/Order",
]].sort_values("Date", ascending=False).reset_index(drop=True)

show["Date"] = show["Date"].dt.strftime("%-m/%-d/%Y")

fmt = {
    "Daily Orders":              "{:,.0f}",
    "Labor Hrs Outbound":        "{:.2f}",
    "Labor Hrs Total":           "{:.2f}",
    "Emp Hrs":                   "{:.2f}",
    "Temp Hrs":                  "{:.2f}",
    "Headcount":                 "{:.0f}",
    "OPLH":                      "{:.1f}",
    "Total Labor Cost/Order":    "${:.2f}",
    "Outbound Labor Cost/Order": "${:.2f}",
    "Avg Shipping Cost/Order":   "${:.2f}",
    "Pkg Cost/Order":            "${:.2f}",
    "Total Cost/Order":          "${:.2f}",
}

st.dataframe(
    show.style.format(fmt, na_rep="—"),
    use_container_width=True,
    height=500,
)

if pd.notna(avg_oplh) and avg_oplh > 0 and pd.notna(avg_labor_cost):
    st.caption(
        f"📊 {len(show):,} days shown  ·  "
        f"Avg OPLH: **{avg_oplh:.1f}**  ·  "
        f"Avg Labor Cost/Order: **${avg_labor_cost:.2f}**  ·  "
        f"Labor rate: **${LABOR_RATE}/hr**  ·  "
        f"Pkg rate: **${PKG_COST_NEW}/order** (from 3/19) / **${PKG_COST}/order** (before)"
    )
else:
    st.caption(f"📊 {len(show):,} days  ·  Labor hours data not available for this range")

# ── Monthly Performance Table ─────────────────────────────────────────────────

st.markdown('<div class="section-header">Monthly Performance Summary</div>',
            unsafe_allow_html=True)

# Build month-level aggregation from full (unfiltered) export + labor data
export_all = export_df.copy()
export_all["Total Ship Cost"] = sum(export_all[c] for c in _SHIP_COLS if c in export_all.columns)
export_all["Pkg Cost"] = export_all["Transaction Date"].apply(
    lambda d: PKG_COST_NEW if d >= PKG_CUTOVER else PKG_COST
)
export_all["Month"] = export_all["Transaction Date"].dt.to_period("M").dt.to_timestamp()

monthly_orders = (
    export_all.groupby("Month")
              .agg(
                  Total_Orders =("OrderID",          "count"),
                  Total_Ship   =("Total Ship Cost",  "sum"),
                  Total_Pkg    =("Pkg Cost",         "sum"),
              )
              .reset_index()
)

if not labor_df.empty:
    labor_all = labor_df.copy()
    labor_all["Month"] = labor_all["Date"].dt.to_period("M").dt.to_timestamp()
    monthly_labor = (
        labor_all.groupby("Month")
                 .agg(Total_Hours=("Total Hours", "sum"))
                 .reset_index()
    )
    monthly_perf = monthly_orders.merge(monthly_labor, on="Month", how="left")
else:
    monthly_perf = monthly_orders.copy()
    monthly_perf["Total_Hours"] = float("nan")

monthly_perf["Labor_Rate"]        = LABOR_RATE
monthly_perf["Total_Pkg_Cost"]    = monthly_perf["Total_Pkg"]
monthly_perf["Pkg_Per_Order"]     = monthly_perf["Total_Pkg_Cost"] / monthly_perf["Total_Orders"]
monthly_perf["Total_Labor_Cost"]  = monthly_perf["Total_Hours"] * LABOR_RATE
monthly_perf["Total_Fulfillment"] = (
    monthly_perf["Total_Labor_Cost"].fillna(0)
    + monthly_perf["Total_Pkg_Cost"]
    + monthly_perf["Total_Ship"].fillna(0)
)

# Grand total row
grand_m = {
    "Month":             "Grand Total",
    "Total_Orders":      monthly_perf["Total_Orders"].sum(),
    "Labor_Rate":        LABOR_RATE,
    "Total_Hours":       monthly_perf["Total_Hours"].sum(),
    "Pkg_Per_Order":     PKG_COST,
    "Total_Pkg_Cost":    monthly_perf["Total_Pkg_Cost"].sum(),
    "Total_Labor_Cost":  monthly_perf["Total_Labor_Cost"].sum(),
    "Total_Ship":        monthly_perf["Total_Ship"].sum(),
    "Total_Fulfillment": monthly_perf["Total_Fulfillment"].sum(),
}
monthly_display = pd.concat([monthly_perf, pd.DataFrame([grand_m])], ignore_index=True)

# Format month column
def _fmt_month(v):
    if isinstance(v, str):
        return v
    try:
        return pd.Timestamp(v).strftime("%-m/1/%Y")
    except Exception:
        return str(v)

monthly_display["Month"] = monthly_display["Month"].apply(_fmt_month)

def _style_grand_m(row):
    if row["Month"] == "Grand Total":
        return ["font-weight: bold; background-color: #f0f4ff"] * len(row)
    return [""] * len(row)

st.dataframe(
    monthly_display.rename(columns={
        "Month":             "Month",
        "Total_Orders":      "Total Orders",
        "Labor_Rate":        "Labor Cost / Hr",
        "Total_Hours":       "Total Labor Hours",
        "Pkg_Per_Order":     "Pkg Cost / Order",
        "Total_Pkg_Cost":    "Total Pkg Cost",
        "Total_Labor_Cost":  "Total Labor Cost",
        "Total_Ship":        "Shipping Cost w/ Surcharges",
        "Total_Fulfillment": "Total Fulfillment Cost",
    }).style
      .format({
          "Total Orders":               "{:,.0f}",
          "Labor Cost / Hr":            "${:.2f}",
          "Total Labor Hours":          "{:,.2f}",
          "Pkg Cost / Order":           "${:.2f}",
          "Total Pkg Cost":             "${:,.2f}",
          "Total Labor Cost":           "${:,.2f}",
          "Shipping Cost w/ Surcharges":"${:,.2f}",
          "Total Fulfillment Cost":     "${:,.2f}",
      }, na_rep="—")
      .apply(_style_grand_m, axis=1),
    use_container_width=True,
    hide_index=True,
    height=min(600, (len(monthly_display) + 1) * 38 + 40),
)

# ── Chart 6: Pre vs Post Negotiation Shipping Cost ────────────────────────────

st.markdown('<div class="section-header">Shipping Cost: Pre vs Post Negotiation</div>',
            unsafe_allow_html=True)

if comparison_df.empty:
    err = _comparison_load_error.get("msg", "")
    if err:
        st.error(f"Could not load comparison data: {err}")
    else:
        st.info("Negotiation comparison data not available. Run `python3 enrich_shipments.py` to generate it.")
else:
    # Apply same date filter as the rest of the dashboard
    cmp_mask = (
        (comparison_df["Transaction Date"].dt.date >= start_date) &
        (comparison_df["Transaction Date"].dt.date <= end_date) &
        (comparison_df["Rate_Lookup_Status"] == "OK")
    )
    cmp = comparison_df[cmp_mask].copy()

    if cmp.empty:
        st.info("No comparison data available for the selected date range.")
    else:
        # Use same period grouping as the rest of the dashboard (weekly or monthly)
        cmp["Period"] = period_start(cmp["Transaction Date"])

        weekly_cmp = (
            cmp.groupby("Period")
               .agg(
                   Shipments        =("OrderID",              "count"),
                   Pre_Neg_Total    =("Pre_Neg_Total_Rate",   "sum"),
                   Post_Neg_Total   =("Current_Total",        "sum"),
                   Total_Savings    =("Savings_Per_Shipment", "sum"),
               )
               .reset_index()
               .sort_values("Period")
        )
        weekly_cmp["Week Label"] = period_label(weekly_cmp["Period"])

        # ── KPI summary row ──────────────────────────────────────────────────
        total_pre  = cmp["Pre_Neg_Total_Rate"].sum()
        total_post = cmp["Current_Total"].sum()
        total_sav  = cmp["Savings_Per_Shipment"].sum()
        pct_saved  = (total_sav / total_pre * 100) if total_pre else 0
        avg_sav_per_ship = total_sav / len(cmp) if len(cmp) else 0

        n1, n2, n3, n4 = st.columns(4)
        n1.metric("Pre-Neg Total Cost",    f"${total_pre:,.0f}")
        n2.metric("Post-Neg Total Cost",   f"${total_post:,.0f}")
        n3.metric("Total Savings",         f"${total_sav:,.0f}",
                  delta=f"{pct_saved:.1f}% vs pre-neg",
                  delta_color="normal")
        n4.metric("Avg Savings / Shipment", f"${avg_sav_per_ship:.2f}")

        st.markdown("")

        # ── Stacked bar (3 segments) ──────────────────────────────────────────
        #
        #  SAVINGS weeks  (post_neg ≤ pre_neg):
        #    ① blue       = post_neg            (actual cost paid)
        #    ② green      = pre_neg − post_neg  (money saved vs old rates)
        #    total bar    = pre_neg
        #
        #  OVERPAY weeks  (post_neg > pre_neg):
        #    ① blue       = pre_neg             (what pre-neg would have cost)
        #    ③ light-red  = post_neg − pre_neg  (extra we pay vs old rates)
        #    total bar    = post_neg
        #
        #  "Pre:" scatter label is pinned at y = pre_neg for every week.
        #    → savings weeks: floats just above the bar top  (pre_neg = bar top)
        #    → overpay weeks: sits at the blue/red boundary  (marks the threshold)

        post_neg_col = weekly_cmp["Post_Neg_Total"].reset_index(drop=True)
        pre_neg_col  = weekly_cmp["Pre_Neg_Total"].reset_index(drop=True)
        savings_raw  = weekly_cmp["Total_Savings"].reset_index(drop=True)

        base_seg    = post_neg_col.clip(upper=pre_neg_col).round(2)   # ① blue
        savings_seg = savings_raw.clip(lower=0).round(2)              # ② green
        premium_seg = (-savings_raw).clip(lower=0).round(2)           # ③ light red

        fig_neg = go.Figure()

        # ① Blue — base cost
        fig_neg.add_bar(
            x=weekly_cmp["Week Label"],
            y=base_seg,
            name="Post-Neg Cost",
            marker_color="#4361ee",
            text=post_neg_col.map("${:,.0f}".format),
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(size=11, color="#ffffff"),
        )

        # ② Green — savings (invisible for overpay weeks)
        fig_neg.add_bar(
            x=weekly_cmp["Week Label"],
            y=savings_seg,
            name="Savings vs Pre-Neg",
            marker_color="#43a878",
            text=savings_seg.map(lambda v: f"+${v:,.0f}" if v > 0 else ""),
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(size=11, color="#ffffff"),
        )

        # ③ Light red — premium / increase over pre-neg (invisible for savings weeks)
        fig_neg.add_bar(
            x=weekly_cmp["Week Label"],
            y=premium_seg,
            name="Increase vs Pre-Neg",
            marker_color="rgba(247, 37, 133, 0.28)",
            marker_line=dict(width=0),
            text=premium_seg.map(lambda v: f"−${v:,.0f}" if v > 0 else ""),
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(size=11, color="#c9184a"),
        )

        # "Pre: $X,XXX" annotations — placed via pixel yshift so they always sit
        # cleanly above the bar top without overlapping bar text.
        # y = top of each bar (max of pre_neg and post_neg); yshift pushes the
        # text 14 px above that in screen space regardless of bar height.
        bar_tops = post_neg_col.clip(lower=pre_neg_col)   # max(post_neg, pre_neg)
        max_bar  = float(bar_tops.max())

        pre_annotations = [
            dict(
                x=row["Week Label"],
                y=float(bar_tops.iloc[i]),
                text=f"Pre: ${row['Pre_Neg_Total']:,.0f}",
                showarrow=False,
                yshift=14,
                font=dict(size=9, color="#6b7a99", family="monospace"),
                xanchor="center",
                yanchor="bottom",
            )
            for i, (_, row) in enumerate(weekly_cmp.iterrows())
        ]

        xaxis_title_neg = "Month" if granularity == "Monthly" else "Week (Monday)"
        fig_neg.update_layout(
            **PLOTLY_THEME,
            barmode="stack",
            xaxis_title=xaxis_title_neg,
            yaxis_title="Total Shipping Cost ($)",
            yaxis_tickprefix="$",
            yaxis_tickformat=",",
            yaxis_range=[0, max_bar * 1.14],   # headroom so labels aren't clipped
            annotations=pre_annotations,
            legend=dict(orientation="h", yanchor="top", y=-0.18, x=0.5, xanchor="center"),
            margin=dict(t=40, b=80, l=10, r=10),
            height=500,
        )
        fig_neg.update_xaxes(type="category")   # prevent "11-03" being parsed as year 2011
        st.plotly_chart(fig_neg, use_container_width=True)

        # ── Carrier breakdown toggle ─────────────────────────────────────────
        with st.expander("📊 Savings breakdown by carrier"):
            if "Carrier" in cmp.columns:
                carrier_cmp = (
                    cmp.groupby("Carrier")
                       .agg(
                           Shipments    =("OrderID",              "count"),
                           Pre_Neg      =("Pre_Neg_Total_Rate",   "sum"),
                           Post_Neg     =("Current_Total",        "sum"),
                           Savings      =("Savings_Per_Shipment", "sum"),
                       )
                       .reset_index()
                       .sort_values("Savings", ascending=False)
                )
                carrier_cmp["Avg Savings / Shipment"] = (
                    carrier_cmp["Savings"] / carrier_cmp["Shipments"]
                ).round(2)
                carrier_cmp["% Saved"] = (
                    carrier_cmp["Savings"] / carrier_cmp["Pre_Neg"] * 100
                ).round(1)

                st.dataframe(
                    carrier_cmp.style.format({
                        "Pre_Neg":               "${:,.2f}",
                        "Post_Neg":              "${:,.2f}",
                        "Savings":               "${:,.2f}",
                        "Avg Savings / Shipment":"${:.2f}",
                        "% Saved":               "{:.1f}%",
                        "Shipments":             "{:,}",
                    }).applymap(
                        lambda v: "color: #43a878; font-weight:600" if isinstance(v, (int, float)) and v > 0
                                  else ("color: #f72585; font-weight:600" if isinstance(v, (int, float)) and v < 0 else ""),
                        subset=["Savings", "Avg Savings / Shipment"]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
                st.caption(
                    "🟢 Green = savings achieved vs pre-negotiation rates  ·  "
                    "🔴 Red = carrier rate is higher post-negotiation (rate table may not apply to this carrier)"
                )
