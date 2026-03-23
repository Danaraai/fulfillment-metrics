"""
dashboard.py  —  OpenStore Operations Dashboard
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

from data_loader import load_export, load_labor_hours, load_daily_metrics

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
COLORS = ["#4361ee", "#f72585", "#4cc9f0", "#f8961e", "#7209b7", "#3a86ff"]

# ── Load data ─────────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def get_data():
    return load_export(), load_labor_hours(), load_daily_metrics()

with st.spinner("Loading data from Google Sheets…"):
    export_df, labor_df, metrics_df = get_data()

if export_df.empty:
    st.error("No shipping data found in the Export tab. Please run the weekly export first.")
    st.stop()

# ── Sidebar filters ───────────────────────────────────────────────────────────

st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c1/Google_"
                 "Sheets_logo_%282014-2020%29.svg/1200px-Google_Sheets_logo_%282014-2020%29.svg.png",
                 width=32)
st.sidebar.title("OpenStore Ops")
st.sidebar.markdown("---")

min_date = export_df["Transaction Date"].min().date()
max_date = export_df["Transaction Date"].max().date()

# Default: show full current year (or all data if older)
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

# Apply date filter
mask = (
    (export_df["Transaction Date"].dt.date >= start_date) &
    (export_df["Transaction Date"].dt.date <= end_date)
)
df = export_df[mask].copy()

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

# ── Header ────────────────────────────────────────────────────────────────────

st.title("📦 OpenStore Operations Dashboard")
st.caption(f"Jack Archer merchant  ·  {start_date.strftime('%b %d, %Y')} – {end_date.strftime('%b %d, %Y')}")

# ── KPI cards ─────────────────────────────────────────────────────────────────

total_orders  = len(df)
avg_ship_cost = df["Original Invoice"].mean() if "Original Invoice" in df.columns else None

# Filter daily metrics to the selected date range
if not metrics_df.empty:
    m_mask = (
        (metrics_df["Date"].dt.date >= start_date) &
        (metrics_df["Date"].dt.date <= end_date)
    )
    mdf = metrics_df[m_mask].copy()
else:
    mdf = pd.DataFrame()

# OPLH and labor cost come straight from the pre-calculated Daily Metrics tab
oplh_col       = "OPLH"
labor_cost_col = next((c for c in (metrics_df.columns if not metrics_df.empty else [])
                       if "total labor cost" in c.lower()), None)

avg_oplh       = mdf[oplh_col].dropna().mean() \
                 if not mdf.empty and oplh_col in mdf.columns else None
avg_labor_cost = mdf[labor_cost_col].dropna().mean() \
                 if not mdf.empty and labor_cost_col else None

k1, k2, k3, k4 = st.columns(4)
kpi_oplh_placeholder       = k3.empty()
kpi_labor_cost_placeholder = k4.empty()
k1.metric("Total Shipments", f"{total_orders:,}")
k2.metric("Avg Shipping Cost / Order",
          f"${avg_ship_cost:.2f}" if pd.notna(avg_ship_cost) and avg_ship_cost != 0 else "—")
# OPLH + Labor Cost KPIs are filled in after the daily table is built below
kpi_oplh_placeholder.metric("Avg OPLH", "—", help="Orders Per Labor Hour (outbound)")
kpi_labor_cost_placeholder.metric("Avg Labor Cost / Order", "—")

# ── Row 2: Carrier mix + Avg cost by carrier ─────────────────────────────────

st.markdown('<div class="section-header">Carrier Performance</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)

if "Carrier" in df.columns:
    carrier_counts = (
        df.groupby("Carrier")
          .size()
          .reset_index(name="Shipments")
          .sort_values("Shipments", ascending=False)
    )

    with c1:
        st.markdown("**Carrier Mix**")
        fig_pie = px.pie(
            carrier_counts,
            names="Carrier",
            values="Shipments",
            color_discrete_sequence=COLORS,
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
        fig_bar = px.bar(
            carrier_cost, y="Carrier", x="Avg Cost ($)",
            orientation="h", color="Avg Cost ($)",
            color_continuous_scale=["#4c9aff", "#f56565"],
            text="Avg Cost ($)",
        )
        fig_bar.update_traces(texttemplate="$%{text:.2f}", textposition="outside")
        fig_bar.update_layout(**PLOTLY_THEME, coloraxis_showscale=False,
                               margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_bar, use_container_width=True)

# ── Row 3: Weekly orders volume ───────────────────────────────────────────────

st.markdown('<div class="section-header">Weekly Volume</div>', unsafe_allow_html=True)

# Use Monday-anchored week start; normalize() strips the sub-second precision
# bug that .start_time produces when using Period("W-SAT")
df["Week"] = (
    df["Transaction Date"]
    - pd.to_timedelta(df["Transaction Date"].dt.dayofweek, unit="D")
).dt.normalize()
weekly_orders = df.groupby("Week").size().reset_index(name="Shipments")

fig_vol = px.line(
    weekly_orders, x="Week", y="Shipments",
    markers=True, color_discrete_sequence=["#4c9aff"],
)
fig_vol.update_layout(**PLOTLY_THEME, margin=dict(t=10, b=10, l=10, r=10))
fig_vol.update_traces(line_width=2, marker_size=5)
st.plotly_chart(fig_vol, use_container_width=True)

# ── Row 4: Cost breakdown per order (weekly) ──────────────────────────────────

if not mdf.empty and "Total Labor Cost Per Order" in mdf.columns:
    st.markdown('<div class="section-header">Cost Per Order Breakdown (Weekly)</div>',
                unsafe_allow_html=True)

    mdf["Week"] = (
        mdf["Date"] - pd.to_timedelta(mdf["Date"].dt.dayofweek, unit="D")
    ).dt.normalize()
    weekly_costs = mdf.groupby("Week").agg(
        Shipping=("Total Labor Cost Per Order", "mean"),  # placeholder until shipping cost is per-order
        Labor=("Total Labor Cost Per Order", "mean"),
        Packaging=("Packaging Cost Per Order", "mean"),
        OPLH=("OPLH", "mean"),
    ).reset_index()

    # Build shipping cost per order from export tab
    weekly_ship = (
        df.groupby("Week")
          .agg(TotalShipping=("Original Invoice", "sum"), Orders=("Original Invoice", "count"))
          .reset_index()
    )
    weekly_ship["Shipping Cost/Order"] = weekly_ship["TotalShipping"] / weekly_ship["Orders"]

    cost_df = weekly_costs.merge(weekly_ship[["Week", "Shipping Cost/Order"]], on="Week", how="left")

    fig_cost = go.Figure()
    if "Shipping Cost/Order" in cost_df.columns:
        fig_cost.add_bar(x=cost_df["Week"], y=cost_df["Shipping Cost/Order"],
                         name="Shipping", marker_color="#4c9aff")
    fig_cost.add_bar(x=cost_df["Week"], y=cost_df["Packaging"],
                     name="Packaging", marker_color="#68d391")
    fig_cost.add_bar(x=cost_df["Week"], y=cost_df["Labor"],
                     name="Labor", marker_color="#f6ad55")
    fig_cost.update_layout(
        **PLOTLY_THEME,
        barmode="stack",
        yaxis_title="Cost per Order ($)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(t=30, b=10, l=10, r=10),
    )
    st.plotly_chart(fig_cost, use_container_width=True)

# ── Row 5: OPLH trend ─────────────────────────────────────────────────────────

if not mdf.empty and "OPLH" in mdf.columns:
    st.markdown('<div class="section-header">Orders Per Labor Hour (OPLH, Total Hours) — Weekly Trend</div>',
                unsafe_allow_html=True)

    mdf["Week"] = (
        mdf["Date"] - pd.to_timedelta(mdf["Date"].dt.dayofweek, unit="D")
    ).dt.normalize()
    oplh_weekly = (
        mdf.dropna(subset=["OPLH"])
           .groupby("Week")["OPLH"]
           .mean()
           .reset_index()
    )

    fig_oplh = px.line(
        oplh_weekly, x="Week", y="OPLH",
        markers=True, color_discrete_sequence=["#68d391"],
    )
    fig_oplh.add_hline(
        y=oplh_weekly["OPLH"].mean(),
        line_dash="dash", line_color="#8a9bb5",
        annotation_text=f"Avg {oplh_weekly['OPLH'].mean():.1f}",
        annotation_position="bottom right",
    )
    fig_oplh.update_layout(**PLOTLY_THEME, margin=dict(t=10, b=10, l=10, r=10))
    fig_oplh.update_traces(line_width=2, marker_size=5)
    st.plotly_chart(fig_oplh, use_container_width=True)

# ── Daily Operations Metrics Table ───────────────────────────────────────────
# Built by aggregating export rows per day, then joining with labor hours.
# OPLH = Daily Orders / Outbound Labor Hours
# Labor Cost/Order = (Total Labor Hours × $18.25/hr) / Daily Orders

st.markdown('<div class="section-header">Daily Operations Metrics</div>', unsafe_allow_html=True)

LABOR_RATE = 18.25  # $/hr — matches your sheet

# ── DEBUG (remove after confirming) ──────────────────────────────────────────
with st.expander("🔍 Debug info (click to expand)", expanded=False):
    st.write(f"**export_df rows:** {len(export_df):,} | **df (filtered) rows:** {len(df):,}")
    st.write(f"**Transaction Date dtype:** `{df['Transaction Date'].dtype}`")
    st.write(f"**Unique dates in df:** {df['Transaction Date'].dt.date.nunique()}")
    st.write(f"**labor_df rows:** {len(labor_df):,} | **ldf rows:** {len(ldf):,}")
    if not ldf.empty:
        st.write(f"**Labor Date dtype:** `{ldf['Date'].dtype}`")
        st.write(f"**Labor unique dates:** {ldf['Date'].dt.date.nunique()}")
    st.write("**First 5 Transaction Dates:**", df["Transaction Date"].head().tolist())
    # Quick groupby test
    test = df["Transaction Date"].dt.date.value_counts().sort_index()
    st.write(f"**Groupby test — unique days:** {len(test)}, **top 5:**", test.tail().to_dict())
# ─────────────────────────────────────────────────────────────────────────────

# Step 1: count orders per calendar day using value_counts (most reliable)
day_counts = (
    df["Transaction Date"].dt.date
      .value_counts()
      .sort_index()
      .reset_index()
)
day_counts.columns = ["Date", "Daily Orders"]
day_counts["Date"] = pd.to_datetime(day_counts["Date"])

if "Original Invoice" in df.columns:
    avg_cost = (
        df.assign(_day=df["Transaction Date"].dt.date)
          .groupby("_day")["Original Invoice"]
          .mean()
          .reset_index()
          .rename(columns={"_day": "Date", "Original Invoice": "Avg_Ship_Cost"})
    )
    avg_cost["Date"] = pd.to_datetime(avg_cost["Date"])
    daily_orders = day_counts.merge(avg_cost, on="Date", how="left")
else:
    daily_orders = day_counts.copy()
    daily_orders["Avg_Ship_Cost"] = float("nan")

# Step 2: join with labor hours — deduplicate labor by date first to prevent row explosion
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

# Step 3: compute OPLH and costs
# OPLH = Daily Orders ÷ Total Labor Hours (all labor, not just outbound)
daily_tbl["OPLH"] = (
    daily_tbl["Daily Orders"] / daily_tbl["Total Hours"]
).where(daily_tbl["Total Hours"].fillna(0) > 0)

daily_tbl["Total Labor Cost/Order ($)"] = (
    daily_tbl["Total Hours"] * LABOR_RATE / daily_tbl["Daily Orders"]
).where(daily_tbl["Daily Orders"] > 0)

daily_tbl["Outbound Labor Cost/Order ($)"] = (
    daily_tbl["Outbound Hours"] * LABOR_RATE / daily_tbl["Daily Orders"]
).where(daily_tbl["Daily Orders"] > 0)

daily_tbl["Avg Shipping Cost/Order ($)"] = daily_tbl["Avg_Ship_Cost"]

# Recalculate KPIs from computed table (overrides stale mdf values)
avg_oplh       = daily_tbl["OPLH"].dropna().mean()
avg_labor_cost = daily_tbl["Total Labor Cost/Order ($)"].dropna().mean()

# Step 4: display
show = daily_tbl.rename(columns={
    "Outbound Hours": "Labor Hrs Outbound",
    "Total Hours":    "Labor Hrs Total",
    "Emp Hours":      "Emp Hrs",
    "Temp Hours":     "Temp Hrs",
})[[
    "Date", "Daily Orders",
    "Labor Hrs Outbound", "Labor Hrs Total", "Emp Hrs", "Temp Hrs", "Headcount",
    "OPLH",
    "Total Labor Cost/Order ($)", "Outbound Labor Cost/Order ($)",
    "Avg Shipping Cost/Order ($)",
]].sort_values("Date", ascending=False).reset_index(drop=True)

show["Date"] = show["Date"].dt.strftime("%-m/%-d/%Y")

fmt = {
    "Daily Orders":                "{:,.0f}",
    "Labor Hrs Outbound":          "{:.2f}",
    "Labor Hrs Total":             "{:.2f}",
    "Emp Hrs":                     "{:.2f}",
    "Temp Hrs":                    "{:.2f}",
    "Headcount":                   "{:.0f}",
    "OPLH":                        "{:.1f}",
    "Total Labor Cost/Order ($)":  "${:.2f}",
    "Outbound Labor Cost/Order ($)": "${:.2f}",
    "Avg Shipping Cost/Order ($)": "${:.2f}",
}

st.dataframe(
    show.style.format(fmt, na_rep="—"),
    use_container_width=True,
    height=500,
)

if pd.notna(avg_oplh) and avg_oplh > 0 and pd.notna(avg_labor_cost):
    st.caption(
        f"📊 {len(show):,} days shown  ·  "
        f"Avg OPLH: **{avg_oplh:.1f}** orders/hr  ·  "
        f"Avg Labor Cost/Order: **${avg_labor_cost:.2f}**  ·  "
        f"Labor rate used: **${LABOR_RATE}/hr**"
    )
else:
    st.caption(f"📊 {len(show):,} days  ·  Labor hours data not available for this range")

# Back-fill KPI cards now that we have computed values
kpi_oplh_placeholder.metric(
    "Avg OPLH",
    f"{avg_oplh:.1f}" if pd.notna(avg_oplh) else "—",
    help="Orders Per Labor Hour (outbound)",
)
kpi_labor_cost_placeholder.metric(
    "Avg Labor Cost / Order",
    f"${avg_labor_cost:.2f}" if pd.notna(avg_labor_cost) else "—",
)
