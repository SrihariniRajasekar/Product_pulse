"""ProductPulse dashboard.  Run:  streamlit run app.py"""
import html
import math
import sqlite3
from datetime import datetime
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parent
DB_PATH = BASE / "data" / "productpulse.db"
REPORT_PATH = BASE / "reports" / "productpulse_summary.xlsx"

YELLOW, PINK, GREEN, BLUE, LAV = "#F5D33F", "#F2B6D8", "#B5C92B", "#7DB7F2", "#B9A6F2"
CORAL, MINT = "#F28B7A", "#3FBF7F"

TYPE_LABELS = {
    "price_high": "Price high",
    "price_low": "Price low",
    "overstock": "Overstock",
    "negative_stock": "Negative stock",
    "stockout_risk": "Stock-out risk",
}
TYPE_COLORS = {"price_high": YELLOW, "price_low": PINK, "overstock": BLUE,
               "negative_stock": GREEN, "stockout_risk": LAV}

st.set_page_config(page_title="ProductPulse", page_icon="◼", layout="wide")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
:root { --glass: rgba(255,255,255,.07); --stroke: rgba(255,255,255,.13); --muted: #9fb6c4; }
html, body, .stApp, [data-testid="stSidebar"], button, input { font-family: 'Plus Jakarta Sans', sans-serif !important; }
.stApp { background:
  radial-gradient(circle at 88% 4%, rgba(245,211,63,.16), transparent 30%),
  radial-gradient(circle at 4% 96%, rgba(125,183,242,.20), transparent 36%),
  linear-gradient(160deg, #0a1b29 0%, #103550 48%, #18586a 100%);
  background-attachment: fixed; }
[data-testid="stAppViewContainer"], header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 1.6rem; max-width: 1320px; }

/* sidebar */
[data-testid="stSidebar"] { background: #0a0a0c; border-right: 1px solid rgba(255,255,255,.06); }
.brand { font-weight: 800; font-size: 1.35rem; letter-spacing: -.02em; }
.brand i { display:inline-block; width:9px; height:9px; border-radius:50%; background:#F2B6D8; margin-left:4px; }
.side-label { font-size:.7rem; letter-spacing:.14em; color:#6f7f8a; text-transform:uppercase; margin:1.4rem 0 .2rem; }
[data-baseweb="tag"] { background:#F5D33F !important; border-radius:999px !important; }
[data-baseweb="tag"] span, [data-baseweb="tag"] svg { color:#111 !important; fill:#111 !important; }

/* header */
.top { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:1.3rem; }
.hello { font-size:2.35rem; font-weight:800; color:#fff; line-height:1.08; letter-spacing:-.03em; }
.hello-sub { color:var(--muted); margin-top:8px; font-size:.95rem; max-width:620px; }
.runpill { background:var(--glass); border:1px solid var(--stroke); border-radius:999px; padding:8px 18px; font-size:.8rem; color:#cfe0ea; white-space:nowrap; }

/* pastel KPI cards */
.kcard { position:relative; overflow:hidden; border-radius:24px; padding:18px 22px; height:138px; color:#15181b; }
.kcard::after { content:""; position:absolute; right:-14px; bottom:-14px; width:92px; height:92px; background:rgba(0,0,0,.09); }
.k-y { background:#F5D33F; } .k-y::after { border-radius:30%; transform:rotate(18deg); }
.k-p { background:#F2B6D8; } .k-p::after { border-radius:50%; }
.k-g { background:#A8BC3A; } .k-g::after { clip-path:polygon(50% 0,100% 100%,0 100%); }
.k-b { background:#AFC8F2; } .k-b::after { clip-path:polygon(50% 0,61% 35%,100% 50%,61% 65%,50% 100%,39% 65%,0 50%,39% 35%); }
.kl { font-weight:700; font-size:.92rem; opacity:.85; }
.kv { font-weight:800; font-size:2.5rem; letter-spacing:-.03em; margin:4px 0 0; line-height:1.1; }
.ks { font-size:.8rem; opacity:.7; font-weight:500; }

/* glass cards */
.card, [class*="st-key-card"] { background:var(--glass); border:1px solid var(--stroke); border-radius:26px;
  padding:20px 24px; backdrop-filter:blur(16px); -webkit-backdrop-filter:blur(16px); box-shadow:0 12px 34px rgba(0,0,0,.2); }
.card { margin-bottom:1rem; }
.h1 { min-height:440px; } .h2 { min-height:420px; }
.card-head { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; font-weight:700; font-size:1.08rem; color:#fff; }
.chip { font-size:.74rem; font-weight:500; color:var(--muted); background:rgba(255,255,255,.08); border-radius:999px; padding:4px 13px; }

/* vertical pill columns */
.pcols { display:flex; gap:14px; align-items:flex-end; }
.pcol { flex:1; display:flex; flex-direction:column; align-items:center; }
.ptrack { width:100%; max-width:46px; height:190px; border-radius:999px; background:rgba(255,255,255,.10); display:flex; align-items:flex-end; overflow:hidden; }
.pfill { width:100%; border-radius:999px; }
.pval { margin-top:10px; font-weight:700; font-size:.95rem; }
.plab { font-size:.74rem; color:var(--muted); margin-top:2px; text-align:center; }

/* donut */
.dwrap { display:flex; align-items:center; gap:22px; flex-wrap:wrap; }
.lg { display:flex; align-items:center; gap:10px; margin:13px 0; }
.dot { width:9px; height:9px; border-radius:50%; flex:none; }
.lg-l { font-size:.78rem; color:var(--muted); }
.lg-v { font-weight:700; font-size:1.15rem; line-height:1.1; }
.lg-s { font-size:.74rem; color:#6fd3a0; margin-left:auto; padding-left:14px; }

/* list rows (challenge style) */
.row { display:flex; align-items:center; gap:13px; padding:11px 0; border-top:1px solid rgba(255,255,255,.08); }
.ring { width:22px; height:22px; border-radius:50%; border:3px solid var(--c); flex:none; }
.rname { flex:1; font-size:.9rem; color:#e9f1f5; }
.rmeta { font-size:.76rem; color:var(--muted); white-space:nowrap; }
.pill { font-size:.74rem; font-weight:700; border-radius:999px; padding:5px 15px; color:#14171a; }
.p-high { background:#F28B7A; } .p-medium { background:#B5C92B; } .p-low { background:#3FBF7F; }

/* horizontal pill bars */
.hrow { display:flex; align-items:center; gap:12px; margin:15px 0; }
.hname { width:36%; font-size:.85rem; color:#dbe8ef; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.htrack { flex:1; height:12px; border-radius:999px; background:rgba(255,255,255,.1); overflow:hidden; }
.hfill { height:100%; border-radius:999px; }
.hval { width:52px; text-align:right; font-weight:700; font-size:.85rem; }

/* detector grid */
.dgrid { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; }
.dbox { background:rgba(255,255,255,.06); border-radius:20px; padding:16px 18px; }
.dnum { font-size:1.75rem; font-weight:800; letter-spacing:-.02em; margin:4px 0 10px; }
.dlab { font-size:.8rem; color:var(--muted); }

/* tabs as pills */
[data-testid="stTabs"] [data-orientation="horizontal"] { border-bottom:0 !important; box-shadow:none !important; }
[data-testid="stTabs"] [role="tablist"] { gap:10px; background:transparent; border:0 !important; box-shadow:none !important; }
[data-testid="stTab"] { background:rgba(255,255,255,.08) !important; border-radius:999px !important; padding:8px 24px !important; height:auto !important; color:#cfe0ea !important; font-weight:600; }
[data-testid="stTab"] p { color:inherit !important; }
[data-testid="stTab"][aria-selected="true"] { background:#F5D33F !important; color:#14171a !important; }
[data-testid="stTab"] .react-aria-SelectionIndicator { display:none !important; }
.stDownloadButton button { border-radius:999px; background:#F5D33F; color:#14171a; border:0; font-weight:700; padding:.4rem 1.3rem; }
@media (max-width: 900px) { .dgrid { grid-template-columns:repeat(2,1fr); } }
</style>
""",
    unsafe_allow_html=True,
)


# ------------------------------------------------------------------ data
@st.cache_data
def load(table: str, mtime: float) -> pd.DataFrame:
    with sqlite3.connect(DB_PATH) as con:
        return pd.read_sql(f"SELECT * FROM {table}", con)


if not DB_PATH.exists():
    st.error("Database not found. Run `python pipeline.py` first.")
    st.stop()

mt = DB_PATH.stat().st_mtime
sku, dq, ev, runlog = (load(t, mt) for t in ["sku_metrics", "dq_issues", "detection_eval", "run_log"])


# ------------------------------------------------------------------ html helpers
def esc(s) -> str:
    return html.escape(str(s))


def card(title, body, chip="", cls="h1") -> str:
    chip_html = f'<span class="chip">{esc(chip)}</span>' if chip else ""
    return f'<div class="card {cls}"><div class="card-head"><span>{esc(title)}</span>{chip_html}</div>{body}</div>'


def kcard(cls, label, value, sub) -> str:
    return (f'<div class="kcard {cls}"><div class="kl">{esc(label)}</div>'
            f'<div class="kv">{value}</div><div class="ks">{esc(sub)}</div></div>')


def pill_columns(items) -> str:
    mx = max([v for _, v, _ in items] + [1])
    cols = ""
    for label, val, color in items:
        h = round(100 * val / mx) if val else 0
        cols += (f'<div class="pcol"><div class="ptrack"><div class="pfill" style="height:{h}%;background:{color}"></div></div>'
                 f'<div class="pval">{val:,}</div><div class="plab">{esc(label)}</div></div>')
    return f'<div class="pcols">{cols}</div>'


def donut(segments, big, small) -> str:
    r, c = 56, 2 * math.pi * 56
    total = sum(v for _, v, _ in segments) or 1
    arcs, off = "", 0.0
    for _, v, color in segments:
        if v:
            length = c * v / total
            arcs += (f'<circle cx="75" cy="75" r="{r}" fill="none" stroke="{color}" stroke-width="17" '
                     f'stroke-dasharray="{length:.2f} {c - length:.2f}" stroke-dashoffset="{-off:.2f}" '
                     f'transform="rotate(-90 75 75)"/>')
            off += length
    svg = (f'<svg width="150" height="150" viewBox="0 0 150 150">'
           f'<circle cx="75" cy="75" r="{r}" fill="none" stroke="rgba(255,255,255,.10)" stroke-width="17"/>{arcs}'
           f'<text x="75" y="76" text-anchor="middle" fill="#fff" font-size="25" font-weight="800">{esc(big)}</text>'
           f'<text x="75" y="95" text-anchor="middle" fill="#9fb6c4" font-size="11">{esc(small)}</text></svg>')
    legend = ""
    for label, v, color in segments:
        share = f"{100 * v / total:.0f}%" if sum(s[1] for s in segments) else "-"
        legend += (f'<div class="lg"><span class="dot" style="background:{color}"></span>'
                   f'<div><div class="lg-l">{esc(label)}</div><div class="lg-v">{v:,}</div></div>'
                   f'<span class="lg-s">{share}</span></div>')
    return f'<div class="dwrap">{svg}<div style="flex:1;min-width:150px">{legend}</div></div>'


def dq_rows(df) -> str:
    if df.empty:
        return '<div class="rmeta">No data-quality issues found.</div>'
    order = {"high": 0, "medium": 1, "low": 2}
    d = df.assign(o=df["severity"].map(order)).sort_values(["o", "records"], ascending=[True, False])
    colors = {"high": CORAL, "medium": GREEN, "low": MINT}
    out = ""
    for _, r in d.head(7).iterrows():
        out += (f'<div class="row"><span class="ring" style="--c:{colors[r.severity]}"></span>'
                f'<span class="rname">{esc(r["check"])}</span>'
                f'<span class="rmeta">{int(r.records):,} &middot; {r.pct}%</span>'
                f'<span class="pill p-{r.severity}">{r.severity.title()}</span></div>')
    if len(d) > 7:
        out += f'<div class="rmeta" style="padding-top:10px">+ {len(d) - 7} more in Explorer</div>'
    return out


def hbars(rows, color, suffix="%") -> str:
    mx = max([v for _, v in rows] + [1])
    out = ""
    for name, v in rows:
        out += (f'<div class="hrow"><span class="hname">{esc(name)}</span>'
                f'<div class="htrack"><div class="hfill" style="width:{100 * v / mx:.0f}%;background:{color}"></div></div>'
                f'<span class="hval">{v:.1f}{suffix}</span></div>')
    return out or '<div class="rmeta">No category in view has 20+ SKUs.</div>'


def style_chart(chart, height=320, grid=False):
    return (chart.properties(height=height)
            .configure(background="transparent")
            .configure_view(stroke=None)
            .configure_axis(grid=grid, gridColor="rgba(255,255,255,0.08)", domain=False, ticks=False,
                            labelColor="#b4c6d2", titleColor="#b4c6d2", labelFontSize=12, titleFontSize=11)
            .configure_legend(orient="top", title=None, labelColor="#dbe8ef", labelFontSize=12))


def head(title, note=""):
    st.markdown(f'<div class="card-head"><span>{esc(title)}</span>'
                f'{f"<span class=chip>{esc(note)}</span>" if note else ""}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.markdown('<div class="brand">productpulse<i></i></div>', unsafe_allow_html=True)
    st.markdown('<div class="side-label">Filters</div>', unsafe_allow_html=True)
    cats = st.multiselect("Category", sorted(sku["category"].unique()), placeholder="All categories")
    types = st.multiselect("Anomaly type", list(TYPE_LABELS), format_func=TYPE_LABELS.get,
                           placeholder="All types")
    high_only = st.toggle("High severity only", help="Both methods agree, or a hard rule was broken.")
    st.markdown('<div class="side-label">Method</div>', unsafe_allow_html=True)
    st.caption("IQR (k = 1.5) and z-score (|z| > 3), per category, on a log scale. "
               "Stock levels are simulated.")

f = sku
if cats:
    f = f[f["category"].isin(cats)]
if types:
    f = f[f["anomaly_type"].apply(lambda s: any(t in s.split(",") for t in types))]
if high_only:
    f = f[f["severity"] == 2]
an = f[f["is_anomaly"] == 1]

# ------------------------------------------------------------------ header
hour = datetime.now().hour
part = "morning" if hour < 12 else "afternoon" if hour < 18 else "evening"
last_run = esc(runlog["run_at"].iloc[-1].replace("T", " "))
st.markdown(
    f'<div class="top"><div><div class="hello">Good {part},<br>here is your catalogue pulse.</div>'
    f'<div class="hello-sub">Pricing and stock anomalies across the Olist product catalogue, '
    f'refreshed by the latest pipeline run.</div></div>'
    f'<div class="runpill">Last run &middot; {last_run}</div></div>',
    unsafe_allow_html=True)

if f.empty:
    st.warning("No SKUs match the current filters.")
    st.stop()

rate = 100 * len(an) / len(f)
k = st.columns(4, gap="medium")
k[0].markdown(kcard("k-y", "SKUs in view", f"{len(f):,}", f"of {len(sku):,} in the catalogue"), unsafe_allow_html=True)
k[1].markdown(kcard("k-p", "Anomalies", f"{len(an):,}", f"{rate:.2f}% of SKUs in view"), unsafe_allow_html=True)
k[2].markdown(kcard("k-g", "Pricing issues", f"{int(f['price_anomaly'].sum()):,}", "price outliers in category"), unsafe_allow_html=True)
k[3].markdown(kcard("k-b", "Stock issues", f"{int(f['stock_anomaly'].sum()):,}", "overstock and stock-outs"), unsafe_allow_html=True)
st.write("")

tab_dash, tab_price, tab_stock, tab_table = st.tabs(["Dashboard", "Pricing", "Stock", "Explorer"])

# ------------------------------------------------------------------ dashboard
with tab_dash:
    a1, a2 = st.columns([3, 2], gap="medium")
    counts = an["anomaly_type"].str.split(",").explode().value_counts()
    items = [(TYPE_LABELS[t], int(counts.get(t, 0)), TYPE_COLORS[t]) for t in TYPE_LABELS]
    a1.markdown(card("Anomalies by type", pill_columns(items), f"{len(an):,} flagged"), unsafe_allow_html=True)

    p_both = int(((an["price_anomaly"] == 1) & (an["stock_anomaly"] == 1)).sum())
    p_only = int(((an["price_anomaly"] == 1) & (an["stock_anomaly"] == 0)).sum())
    s_only = int(((an["price_anomaly"] == 0) & (an["stock_anomaly"] == 1)).sum())
    a2.markdown(card("Overview", donut(
        [("Pricing only", p_only, YELLOW), ("Stock only", s_only, BLUE), ("Pricing and stock", p_both, PINK)],
        f"{rate:.1f}%", "anomaly rate"), "flagged SKUs"), unsafe_allow_html=True)

    b1, b2 = st.columns([3, 2], gap="medium")
    b1.markdown(card("Data quality checks", dq_rows(dq), f"{len(dq)} findings", "h2"), unsafe_allow_html=True)
    g = f.groupby("category").agg(skus=("product_id", "size"), rate=("is_anomaly", "mean")).reset_index()
    g = g[g["skus"] >= 20].nlargest(7, "rate")
    b2.markdown(card("Hotspot categories", hbars([(r.category, r.rate * 100) for r in g.itertuples()], YELLOW),
                     "20+ SKUs", "h2"), unsafe_allow_html=True)

    names = {"price_spike": "Price spikes", "price_drop": "Price drops",
             "negative_stock": "Negative stock", "overstock": "Overstock"}
    boxes = ""
    for r in ev.itertuples():
        v = 0.0 if pd.isna(r.recall_pct) else float(r.recall_pct)
        boxes += (f'<div class="dbox"><div class="dlab">{esc(names.get(r.injected_type, r.injected_type))}</div>'
                  f'<div class="dnum">{v:.1f}%</div>'
                  f'<div class="htrack"><div class="hfill" style="width:{v}%;background:{MINT}"></div></div>'
                  f'<div class="dlab" style="margin-top:8px">{int(r.detected)} of {int(r.injected)} caught</div></div>')
    st.markdown(card("Detector check", f'<div class="dgrid">{boxes}</div>',
                     "recall on injected anomalies", ""), unsafe_allow_html=True)

# ------------------------------------------------------------------ pricing
with tab_price:
    with st.container(key="card_scatter"):
        head("Price against demand", "log scales, normal SKUs sampled to 4,000")
        p = f[f["current_price"].notna()].copy()
        p["status"] = np.where(p["price_anomaly"] == 1, "Anomaly", "Normal")
        p["units_sold"] = p["units_sold"].clip(lower=1)
        normal = p[p["status"] == "Normal"]
        plot = pd.concat([p[p["status"] == "Anomaly"], normal.sample(min(4000, len(normal)), random_state=1)])
        pts = alt.Chart(plot).mark_circle(size=48).encode(
            x=alt.X("units_sold:Q", scale=alt.Scale(type="log"), title="Units sold"),
            y=alt.Y("current_price:Q", scale=alt.Scale(type="log"), title="Price (R$)"),
            color=alt.Color("status:N", scale=alt.Scale(domain=["Normal", "Anomaly"], range=["#6f93a6", YELLOW])),
            opacity=alt.condition(alt.datum.status == "Anomaly", alt.value(0.95), alt.value(0.40)),
            tooltip=["product_id", "category", "current_price", "units_sold", "price_z"],
        )
        st.altair_chart(style_chart(pts, 380, grid=True), width="stretch")

    st.write("")
    with st.container(key="card_worst"):
        head("Most extreme price outliers", "top 20 by |z|")
        worst = (p[p["price_anomaly"] == 1].assign(z=lambda d: d["price_z"].abs())
                 .nlargest(20, "z")[["product_id", "category", "current_price", "units_sold", "price_z"]])
        st.dataframe(worst, hide_index=True, width="stretch", column_config={
            "product_id": "SKU", "category": "Category",
            "current_price": st.column_config.NumberColumn("Price", format="R$ %.2f"),
            "units_sold": "Units sold", "price_z": st.column_config.NumberColumn("z-score", format="%.2f"),
        })

# ------------------------------------------------------------------ stock
with tab_stock:
    s = st.columns(3, gap="medium")
    s[0].markdown(kcard("k-p", "Negative stock", f"{int((f['stock_level'] < 0).sum()):,}", "impossible values"), unsafe_allow_html=True)
    s[1].markdown(kcard("k-b", "Overstock", f"{int(f['anomaly_type'].str.contains('overstock').sum()):,}", "cover far above category norm"), unsafe_allow_html=True)
    s[2].markdown(kcard("k-y", "Stock-out risk", f"{int(f['anomaly_type'].str.contains('stockout_risk').sum()):,}", "zero stock on top sellers"), unsafe_allow_html=True)
    st.write("")
    with st.container(key="card_hist"):
        head("Stock cover", "units in stock per unit sold, trimmed at the 99th percentile")
        trimmed = f[f["stock_cover"] <= f["stock_cover"].quantile(0.99)]
        hist = alt.Chart(trimmed).mark_bar(color=BLUE, opacity=.9, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            x=alt.X("stock_cover:Q", bin=alt.Bin(maxbins=45), title="Stock cover"),
            y=alt.Y("count()", title=None))
        st.altair_chart(style_chart(hist, 320), width="stretch")

# ------------------------------------------------------------------ explorer
with tab_table:
    with st.container(key="card_explorer"):
        head("Flagged SKUs", f"{len(an):,} rows, highest severity first")
        cols = ["product_id", "category", "units_sold", "current_price", "stock_level",
                "stock_cover", "price_z", "stock_z", "anomaly_type", "severity"]
        out = an.sort_values(["severity", "price_z"], ascending=[False, False])[cols].copy()
        out["anomaly_type"] = out["anomaly_type"].apply(
            lambda s_: ", ".join(TYPE_LABELS.get(t, t) for t in s_.split(",")))
        out["severity"] = out["severity"].map({2: "High", 1: "Medium"})
        st.dataframe(out, hide_index=True, width="stretch", height=460, column_config={
            "product_id": "SKU", "category": "Category", "units_sold": "Units sold",
            "current_price": st.column_config.NumberColumn("Price", format="R$ %.2f"),
            "stock_level": "Stock", "stock_cover": st.column_config.NumberColumn("Cover", format="%.1f"),
            "price_z": st.column_config.NumberColumn("Price z", format="%.2f"),
            "stock_z": st.column_config.NumberColumn("Stock z", format="%.2f"),
            "anomaly_type": "Type", "severity": "Severity",
        })
        d1, d2, _ = st.columns([1, 1, 3])
        d1.download_button("Download CSV", out.to_csv(index=False).encode(), "anomalies.csv", "text/csv")
        if REPORT_PATH.exists():
            d2.download_button("Excel report", REPORT_PATH.read_bytes(), REPORT_PATH.name,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.write("")
    with st.container(key="card_dqtable"):
        head("Data-quality findings", "raw data, before cleaning")
        st.dataframe(dq.rename(columns=str.title), hide_index=True, width="stretch")
