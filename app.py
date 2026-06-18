"""
Home Leads Funnel — automated dashboard.

Run:  streamlit run app.py
Then drag in an updated "Home Leads Funnel" workbook (same sheet layout).
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import data as D
import report

st.set_page_config(page_title="Home Leads Funnel", page_icon="📊", layout="wide")

def _theme_type() -> str:
    """Detect the active Streamlit theme ('dark' or 'light'), robust across versions."""
    try:
        t = getattr(st.context, "theme", None)
        if t and getattr(t, "type", None):
            return t.type
    except Exception:
        pass
    return st.get_option("theme.base") or "light"


IS_DARK = _theme_type() == "dark"
PLOTLY_TMPL = "plotly_dark" if IS_DARK else "plotly_white"
LINE_COLOR = "#e2e8f0" if IS_DARK else "#0f172a"   # conversion line, readable on either bg

# --- theme-adaptive styling (colours chosen per active theme) ------------
CARD_BG = "#1f2430" if IS_DARK else "#ffffff"
CARD_BORDER = "rgba(255,255,255,.12)" if IS_DARK else "#e7e9ee"
CARD_SHADOW = "0 1px 4px rgba(0,0,0,.35)" if IS_DARK else "0 1px 4px rgba(15,23,42,.08)"
st.markdown(
    f"""
    <style>
      .block-container {{padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1500px;}}
      /* Cards use a theme-matched panel colour; text keeps the theme's own
         (light) colour, so values stay legible in both dark and light mode. */
      div[data-testid="stMetric"] {{
          background: {CARD_BG};
          border: 1px solid {CARD_BORDER};
          border-radius: 14px; padding: 14px 18px;
          box-shadow: {CARD_SHADOW};
      }}
      div[data-testid="stMetricValue"] {{font-size: 1.45rem; white-space: nowrap;}}
      /* Chart cards: transparent charts sit on the same panel colour. */
      div[data-testid="stPlotlyChart"], .stPlotlyChart {{
          background: {CARD_BG};
          border: 1px solid {CARD_BORDER};
          border-radius: 14px; padding: 6px;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


def _themed(fig):
    """Apply theme template + transparent backgrounds so charts blend with the card."""
    fig.update_layout(
        template=PLOTLY_TMPL,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    return fig


@st.cache_data(show_spinner="Reading workbook…")
def _load(file_bytes: bytes):
    # Cached on the file's *content*, so a new/edited upload always reloads
    # (even if its name and size happen to match a previous one).
    return D.load_workbook(file_bytes)


def _money(v: float) -> str:
    """Compact currency so it never overflows the KPI card (e.g. AED 2.80M)."""
    v = float(v or 0)
    if abs(v) >= 1e6:
        return f"AED {v/1e6:.2f}M"
    if abs(v) >= 1e3:
        return f"AED {v/1e3:.0f}K"
    return f"AED {v:,.0f}"


def kpi_row(funnel: dict, df: pd.DataFrame):
    total = funnel["Total Leads"] or 1
    done = funnel["Deals Done"]
    deal_value = df.loc[df["Status"] == "Done", "_deal_value"].sum()
    c = st.columns(6)
    c[0].metric("Total Leads", f"{funnel['Total Leads']:,}")
    c[1].metric("Deals Done", f"{done:,}", f"{done/total*100:.1f}% conversion")
    # No leading "-" => up arrow; inverse paints it red (high rejection = bad).
    c[2].metric("Rejected", f"{funnel['Rejected']:,}", f"{funnel['Rejected']/total*100:.1f}% of leads", delta_color="inverse")
    c[3].metric("Pending", f"{funnel['Pending']:,}", f"{funnel['Pending']/total*100:.1f}% of leads", delta_color="off")
    c[4].metric("Awaiting Conf.", f"{funnel['Awaiting Confirmation']:,}")
    c[5].metric("Value Traded (Done)", _money(deal_value), help=f"AED {deal_value:,.0f}")


def funnel_chart(funnel: dict):
    stages = ["Total Leads", "Deals Done"]
    fig = go.Figure(go.Funnel(
        y=stages,
        x=[funnel["Total Leads"], funnel["Deals Done"]],
        textinfo="value+percent initial",
        marker={"color": ["#6366f1", "#16a34a"]},
    ))
    fig.update_layout(height=300, margin=dict(l=10, r=10, t=30, b=10),
                      title="Lead → Deal conversion")
    return _themed(fig)


def status_donut(funnel: dict):
    labels = D.STATUS_ORDER
    vals = [funnel["Deals Done"] if s == "Done" else funnel[s] for s in labels]
    fig = go.Figure(go.Pie(
        labels=labels, values=vals, hole=.6,
        marker=dict(colors=[D.STATUS_COLORS[s] for s in labels]),
        sort=False,
    ))
    fig.update_layout(height=300, margin=dict(l=10, r=10, t=30, b=10),
                      title="Outcome split", legend=dict(orientation="h", y=-0.1))
    return _themed(fig)


def channel_bar(summ: pd.DataFrame):
    d = summ[summ["Channel"] != "ALL CHANNELS"].copy()
    long = d.melt(id_vars="Channel", value_vars=D.STATUS_ORDER, var_name="Status", value_name="Count")
    fig = px.bar(long, x="Count", y="Channel", color="Status", orientation="h",
                 color_discrete_map=D.STATUS_COLORS,
                 category_orders={"Status": D.STATUS_ORDER})
    fig.update_layout(height=460, barmode="stack",
                      margin=dict(l=10, r=10, t=40, b=10),
                      yaxis=dict(categoryorder="total ascending"),
                      title="Funnel by channel", legend=dict(orientation="h", y=-0.12))
    return _themed(fig)


def trend_chart(tr: pd.DataFrame, bucket: str):
    long = tr.melt(id_vars="label", value_vars=D.STATUS_ORDER, var_name="Status", value_name="Count")
    fig = px.bar(long, x="label", y="Count", color="Status",
                 color_discrete_map=D.STATUS_COLORS,
                 category_orders={"Status": D.STATUS_ORDER, "label": list(tr["label"])})
    fig.add_trace(go.Scatter(
        x=tr["label"], y=tr["Conversion %"], name="Conversion %",
        mode="lines+markers", yaxis="y2", line=dict(color=LINE_COLOR, width=2)))
    fig.update_layout(
        height=420, barmode="stack",
        margin=dict(l=10, r=10, t=40, b=10),
        title=f"Leads & conversion over time — {bucket}",
        yaxis=dict(title="Leads"),
        yaxis2=dict(title="Conversion %", overlaying="y", side="right", range=[0, 100], showgrid=False),
        legend=dict(orientation="h", y=-0.2),
    )
    return _themed(fig)


# =========================================================================
# Sidebar — input & filters
# =========================================================================
st.sidebar.title("📊 Home Leads Funnel")
up = st.sidebar.file_uploader("Upload updated workbook (.xlsx)", type=["xlsx"])

if up is None:
    st.title("Home Leads Funnel — automated dashboard")
    st.info("⬅️ Upload a **Home Leads Funnel** workbook to begin. "
            "It should contain one sheet per channel (the same layout as the source file).")
    st.stop()

try:
    df_all = _load(up.getvalue())
except Exception as e:
    st.error(f"Could not read **{up.name}**: {e}")
    st.stop()

st.sidebar.caption(f"Loaded: **{up.name}** — {len(df_all):,} deals · {df_all['Channel'].nunique()} channels")

# Filters
channels = sorted(df_all["Channel"].unique())
sel_channels = st.sidebar.multiselect("Channels", channels, default=channels)

dmin, dmax = df_all["_date"].min(), df_all["_date"].max()
if pd.notna(dmin) and pd.notna(dmax):
    dr = st.sidebar.date_input("Deal date range", value=(dmin.date(), dmax.date()),
                               min_value=dmin.date(), max_value=dmax.date())
else:
    dr = None

bucket = st.sidebar.radio("Time bucket", D.BUCKET_OPTIONS, index=2)  # Monthly default

# Apply filters
df = df_all[df_all["Channel"].isin(sel_channels)].copy()
if dr and isinstance(dr, (list, tuple)) and len(dr) == 2:
    lo, hi = pd.Timestamp(dr[0]), pd.Timestamp(dr[1]) + pd.Timedelta(days=1)
    df = df[(df["_date"] >= lo) & (df["_date"] < hi)]

if df.empty:
    st.warning("No deals match the current filters.")
    st.stop()

# =========================================================================
# Main
# =========================================================================
period_txt = ""
if pd.notna(df["_date"].min()):
    period_txt = f" · {df['_date'].min():%d %b %Y} → {df['_date'].max():%d %b %Y}"
st.title("Home Leads Funnel — automated dashboard")
st.caption(f"{len(df):,} deals · {len(sel_channels)} channel(s){period_txt}")

funnel = D.funnel_counts(df)
kpi_row(funnel, df)
st.divider()

c1, c2 = st.columns([1, 1])
c1.plotly_chart(funnel_chart(funnel), use_container_width=True)
c2.plotly_chart(status_donut(funnel), use_container_width=True)

st.plotly_chart(trend_chart(D.trend(df, bucket), bucket), use_container_width=True)

st.subheader("By channel")
summ = D.channel_summary(df)
cc1, cc2 = st.columns([1.1, 1])
cc1.plotly_chart(channel_bar(summ), use_container_width=True)
disp = summ.rename(columns={"Done": "Deals Done"})
for col in ["Total Leads", "Deals Done", "Rejected", "Pending", "Awaiting Confirmation"]:
    disp[col] = disp[col].astype(int)
cc2.dataframe(
    disp.style.format({"Conversion %": "{:.1f}%"})
        .background_gradient(subset=["Conversion %"], cmap="RdYlGn", vmin=0, vmax=100),
    width="stretch", height=460, hide_index=True,
)

st.subheader("Why leads drop off & what's being traded")
b1, b2, b3 = st.columns(3)
rr = D.reject_reasons(df, 10)
if not rr.empty:
    fig = px.bar(rr.sort_values("Count"), x="Count", y="Reason", orientation="h",
                 color_discrete_sequence=["#dc2626"])
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=40, b=10),
                      title="Top rejection reasons", yaxis_title=None)
    b1.plotly_chart(_themed(fig), use_container_width=True)

cat = D.breakdown(df, "Category", 8)
if not cat.empty:
    fig = px.pie(cat, names="Category", values="Count", hole=.5)
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=40, b=10),
                      title="Deals by category", legend=dict(orientation="h", y=-0.15))
    b2.plotly_chart(_themed(fig), use_container_width=True)

brand = D.breakdown(df, "Brand", 8)
if not brand.empty:
    fig = px.bar(brand.sort_values("Count"), x="Count", y="Brand", orientation="h",
                 color_discrete_sequence=["#6366f1"])
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=40, b=10),
                      title="Top brands", yaxis_title=None)
    b3.plotly_chart(_themed(fig), use_container_width=True)

# Downloads & detail
st.divider()
st.subheader("Export & detail")
tr = D.trend(df, bucket)
st.download_button(
    "⬇️ Download summary report (.xlsx)",
    data=report.build_report_xlsx(df, bucket),
    file_name="Home_Leads_Funnel_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Channel-wise distribution with counts + % side by side, plus charts.",
)
with st.expander("Per-channel summary table (also in the export)"):
    st.dataframe(disp, width="stretch", hide_index=True)
with st.expander(f"Time-bucket table — {bucket}"):
    st.dataframe(tr[["label"] + D.STATUS_ORDER + ["Total Leads", "Conversion %"]],
                 width="stretch", hide_index=True)
