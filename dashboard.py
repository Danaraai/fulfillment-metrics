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

from data_loader import load_export, load_daily_metrics

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
  [data-testid="metric-container"] {
    background: #1e2130;
    border: 1px solid #2d3250;
    border-radius: 10px;
    padding: 16px 20px;
  }
  [data-testid="metric-container"] label { color: #8a9bb5 !important; font-size: 13px; }
  [data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 28px; font-weight: 700; color: #ffffff;
  }
  .section-header {
    font-size: 16px; font-weight: 600; color: #c5d0e6;
    margin: 24px 0 8px; padding-bottom: 6px;
    border-bottom: 1px solid #2d3250;
  }
  .stSidebar { background: #151827; }
</style>
""", unsafe_allow_html=True)

PLOTLY_THEME = dict(
    plot_bgcolor="#1e2130",
    paper_bgcolor="#151827",
    font_color="#c5d0e6",
    xaxis=dict(gridcolor="#2d3250", linecolor="#2d3250"),
    yaxis=dict(gridcolor="#2d3250", linecolor="#2d3250"),
)
COLORS = ["#4c9aff", "#f56565", "#68d391", "#f6ad55", "#b794f4", "#76e4f7"]

# ── Load data ─────────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def get_data():
    return load_export(), load_daily_metrics()

with st.spinner("Loading data from Google Sheets…"):
    export_df, metrics_df = get_data()

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

if not metrics_df.empty:
    m_mask = (
        (metrics_df["Date"].dt.date >= start_date) &
        (metrics_df["Date"].dt.date <= end_date)
    )
    mdf = metrics_df[m_mask].copy()
else:
    mdf = pd.DataFrame()

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

total_orders = len(df)
avg_ship_cost = df["Original Invoice"].mean() if "Original Invoice" in df.columns else 0

# OPLH and labor cost from metrics tab
avg_oplh = mdf["OPLH"].dropna().mean() if not mdf.empty and "OPLH" in mdf.columns else None
avg_labor_cost = (
    mdf["Total Labor Cost Per Order"].dropna().mean()
    if not mdf.empty and "Total Labor Cost Per Order" in mdf.columns else None
)
avg_pkg_cost = (
    mdf["Packaging Cost Per Order"].dropna().mean()
    if not mdf.empty and "Packaging Cost Per Order" in mdf.columns else None
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Shipments", f"{total_orders:,}")
k2.metric("Avg Shipping Cost / Order", f"${avg_ship_cost:.2f}" if pd.notna(avg_ship_cost) and avg_ship_cost != 0 else "—")
k3.metric("Avg OPLH", f"{avg_oplh:.1f}" if avg_oplh and avg_oplh == avg_oplh else "—",
          help="Orders Per Labor Hour (outbound)")
k4.metric("Avg Labor Cost / Order", f"${avg_labor_cost:.2f}" if avg_labor_cost and avg_labor_cost == avg_labor_cost else "—")

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

df["Week"] = df["Transaction Date"].dt.to_period("W-SAT").apply(lambda r: r.start_time)
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

    mdf["Week"] = mdf["Date"].dt.to_period("W-SAT").apply(lambda r: r.start_time)
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
    st.markdown('<div class="section-header">Orders Per Labor Hour (OPLH) — Weekly Trend</div>',
                unsafe_allow_html=True)

    mdf["Week"] = mdf["Date"].dt.to_period("W-SAT").apply(lambda r: r.start_time)
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

# ── Row 6: Data table ─────────────────────────────────────────────────────────

with st.expander("View raw daily metrics table"):
    if not mdf.empty:
        display_cols = [c for c in [
            "Date", "Daily Orders", "Labor Hours Outbound", "Labor Hours Total",
            "OPLH", "Total Labor Cost Per Order", "Outbound Labor Cost Per Order",
            "Packaging Cost Per Order"
        ] if c in mdf.columns]
        st.dataframe(
            mdf[display_cols].sort_values("Date", ascending=False).reset_index(drop=True),
            use_container_width=True,
            height=400,
        )
    else:
        st.info("Daily Metrics tab not yet available. Run the Apps Script first.")
