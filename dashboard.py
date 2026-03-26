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

from data_loader import load_export, load_labor_hours, load_daily_metrics, load_comparison, _comparison_load_error

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="OpenStore Operations",
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
PKG_COST     = 0.31    # $/order (fixed)

# ── Load data ─────────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def get_data():
    return load_export(), load_labor_hours(), load_daily_metrics()

with st.spinner("Loading data from Google Sheets…"):
    export_df, labor_df, metrics_df = get_data()

with st.spinner("Loading negotiation comparison data…"):
    comparison_df = load_comparison()

if export_df.empty:
    st.error("No shipping data found in the Export tab. Please run the weekly export first.")
    st.stop()

# ── Sidebar filters ───────────────────────────────────────────────────────────

st.sidebar.title("📦 OpenStore Ops")
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

# Apply date filter to export
mask = (
    (export_df["Transaction Date"].dt.date >= start_date) &
    (export_df["Transaction Date"].dt.date <= end_date)
)
df = export_df[mask].copy()

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

# ── Week helper ──────────────────────────────────────────────────────────────
# Monday-anchored week start, normalized (no time component)

def week_start(series: pd.Series) -> pd.Series:
    return (series - pd.to_timedelta(series.dt.dayofweek, unit="D")).dt.normalize()

df["Week"] = week_start(df["Transaction Date"])

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

# Daily avg shipping cost
if "Original Invoice" in df.columns:
    avg_cost_daily = (
        df.assign(_day=df["Transaction Date"].dt.date)
          .groupby("_day")["Original Invoice"]
          .mean()
          .reset_index()
          .rename(columns={"_day": "Date", "Original Invoice": "Avg_Ship_Cost"})
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
daily_tbl["Pkg Cost/Order"] = PKG_COST

daily_tbl["Total Cost/Order"] = (
    daily_tbl["Total Labor Cost/Order"].fillna(0)
    + PKG_COST
    + daily_tbl["Avg Shipping Cost/Order"].fillna(0)
)

# ── Build WEEKLY aggregation ──────────────────────────────────────────────────

# Weekly orders + shipping from export
weekly_export = (
    df.groupby("Week")
      .agg(
          Orders=("Transaction Date", "count"),
          Avg_Ship=("Original Invoice", "mean"),
      )
      .reset_index()
)

# Weekly labor from ldf
if not ldf.empty:
    ldf2 = ldf.copy()
    ldf2["Week"] = week_start(ldf2["Date"])
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
weekly["Pkg Cost/Order"]            = PKG_COST
weekly["Total Cost/Order"]          = (
    weekly["Total Labor Cost/Order"].fillna(0)
    + PKG_COST
    + weekly["Avg_Ship"].fillna(0)
)
weekly["Week Label"] = weekly["Week"].dt.strftime("%-m-%d")

# ── KPI cards ─────────────────────────────────────────────────────────────────

total_orders  = len(df)
avg_ship_cost = df["Original Invoice"].mean() if "Original Invoice" in df.columns else None
avg_oplh      = daily_tbl["OPLH"].dropna().mean()
avg_labor_cost = daily_tbl["Total Labor Cost/Order"].dropna().mean()

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Shipments", f"{total_orders:,}")
k2.metric("Avg Shipping Cost / Order",
          f"${avg_ship_cost:.2f}" if pd.notna(avg_ship_cost) and avg_ship_cost != 0 else "—")
k3.metric("Avg OPLH (Total Hours)",
          f"{avg_oplh:.1f}" if pd.notna(avg_oplh) else "—",
          help="Orders Per Total Labor Hour")
k4.metric("Avg Labor Cost / Order",
          f"${avg_labor_cost:.2f}" if pd.notna(avg_labor_cost) else "—")

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

# Grand total / average row
grand = {
    "Week": "Grand Total",
    "Avg OPLH":                         tbl["Avg OPLH"].mean(),
    "Total Labor Cost/Order":           tbl["Total Labor Cost/Order"].mean(),
    "Packaging Cost/Order":             PKG_COST,
    "Outbound Labor Cost/Order":        tbl["Outbound Labor Cost/Order"].mean(),
    "Avg Shipping Cost w/ Surcharges":  tbl["Avg Shipping Cost w/ Surcharges"].mean(),
    "Cost per Order (Pick, Pack, Ship)": tbl["Cost per Order (Pick, Pack, Ship)"].mean(),
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
    transit["Week Label"] = transit["Week"].dt.strftime("%-m-%d")
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
        f"Pkg rate: **${PKG_COST}/order**"
    )
else:
    st.caption(f"📊 {len(show):,} days  ·  Labor hours data not available for this range")

# ── Chart 6: Pre vs Post Negotiation Shipping Cost (Stacked Weekly Bar) ────────

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
        # Week aggregation (Monday-anchored, same as rest of dashboard)
        cmp["Week"] = week_start(cmp["Transaction Date"])

        weekly_cmp = (
            cmp.groupby("Week")
               .agg(
                   Shipments        =("OrderID",              "count"),
                   Pre_Neg_Total    =("Pre_Neg_Total_Rate",   "sum"),
                   Post_Neg_Total   =("Current_Total",        "sum"),
                   Total_Savings    =("Savings_Per_Shipment", "sum"),
               )
               .reset_index()
               .sort_values("Week")
        )
        weekly_cmp["Week Label"] = weekly_cmp["Week"].dt.strftime("%-m-%d")

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

        post_neg_col = weekly_cmp["Post_Neg_Total"]
        pre_neg_col  = weekly_cmp["Pre_Neg_Total"]
        savings_raw  = weekly_cmp["Total_Savings"]

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
            marker_line=dict(color="#f72585", width=1),
            text=premium_seg.map(lambda v: f"−${v:,.0f}" if v > 0 else ""),
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(size=11, color="#c9184a"),
        )

        # Pre-neg total shown as hover only — avoids label overlap on dense weeks
        fig_neg.add_scatter(
            x=weekly_cmp["Week Label"],
            y=pre_neg_col,
            mode="markers",
            marker=dict(size=0, opacity=0),
            customdata=pre_neg_col.values,
            hovertemplate="Pre-Neg: $%{customdata:,.0f}<extra></extra>",
            showlegend=False,
        )

        fig_neg.update_layout(
            **PLOTLY_THEME,
            barmode="stack",
            xaxis_title="Week (Monday)",
            yaxis_title="Total Shipping Cost ($)",
            yaxis_tickprefix="$",
            yaxis_tickformat=",",
            legend=dict(orientation="h", yanchor="top", y=-0.18, x=0.5, xanchor="center"),
            margin=dict(t=40, b=80, l=10, r=10),
            height=460,
        )
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
