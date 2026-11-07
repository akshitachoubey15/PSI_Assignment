"""TPM Status Dashboard – Streamlit app.  Run: streamlit run app.py"""
import io
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

from report_generator import (TrackerError, build_reports, detect_as_of, load_tracker, normalized_items,
                              render_html, render_markdown)

SAMPLE = Path(__file__).parent / "sample_data" / "sample_tracker.xlsx"
RAG = {"Red": "#c62828", "Amber": "#ef8f00", "Green": "#2e7d32"}
LEVEL = {"Low": 1, "Medium": 2, "High": 3}


# Works on old and new Streamlit versions (the width / iframe APIs changed in 2025-26).
def stretch(fn, *a, **k):
    try:
        return fn(*a, width="stretch", **k)
    except TypeError:
        return fn(*a, use_container_width=True, **k)


def embed_html(page, height):
    if hasattr(st, "iframe"):
        return st.iframe(page, height=height)
    return components.html(page, height=height, scrolling=True)


st.set_page_config(page_title="TPM Status Dashboard", page_icon="📊", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1.4rem;max-width:1250px}
.hero{background:linear-gradient(90deg,#0b2545,#13315c);color:#fff;border-radius:12px;padding:18px 24px;margin-bottom:14px}
.hero h1{margin:0;font-size:26px;color:#fff}.hero .s{opacity:.8;font-size:13px;margin-top:2px}
.kpi{background:#fff;border:1px solid #e4e7eb;border-radius:12px;padding:12px 16px;box-shadow:0 1px 2px rgba(0,0,0,.05)}
.kpi .n{font-size:30px;font-weight:700;line-height:1.1}.kpi .l{color:#52606d;font-size:12.5px}
.pcard{background:#fff;border:1px solid #e4e7eb;border-radius:12px;padding:14px 18px;height:100%;box-shadow:0 1px 2px rgba(0,0,0,.05)}
.pcard h3{margin:0;font-size:18px;color:#0b2545}.pcard .c{color:#52606d;font-size:13px}
.badge{color:#fff;border-radius:14px;padding:2px 12px;font-weight:700;font-size:12px;float:right}
.pcard .row{margin-top:8px;font-size:13px}.pcard .k{color:#7b8794}
</style>""", unsafe_allow_html=True)

# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.header("1. Your tracker")
    uploaded = st.file_uploader("Excel (.xlsx) or CSV", type=["xlsx", "csv"], accept_multiple_files=True,
                                help="One workbook with Summary and Risks sheets, or two CSV files (Summary + Risks).")
    use_sample = st.toggle("Use the sample tracker", value=False, disabled=bool(uploaded),
                           help="Switched off automatically when you upload a file.")
    st.header("2. Report date")
    as_of_override = st.date_input("Report as of", value=None, help="Empty = date written in the tracker, else today.")
    st.caption("Overdue and 'next 7 days' are counted from this date.")

files = []
if uploaded:
    files = [(f.name, f.getvalue()) for f in uploaded]
elif use_sample:
    files = [(SAMPLE.name, SAMPLE.read_bytes())]

if not files:
    st.markdown('<div class="hero"><h1>📊 TPM Status Dashboard</h1><div class="s">Upload a project tracker to see portfolio health, risks and ready-to-send client reports.</div></div>', unsafe_allow_html=True)
    st.info("Upload your tracker in the sidebar, or switch on **Use the sample tracker** to see an example.")
    st.stop()

try:
    tracker = load_tracker(files)
    as_of = as_of_override or detect_as_of(tracker) or date.today()
    reports = build_reports(tracker, as_of)
    items = pd.DataFrame(normalized_items(tracker))
except TrackerError as err:
    st.error(str(err))
    st.stop()
except Exception as err:
    st.error(f"Something unexpected went wrong while reading the tracker ({type(err).__name__}). Compare your file with the sample.")
    st.stop()

with st.sidebar:
    st.header("3. Filter")
    names = [r.project for r in reports]
    chosen = st.multiselect("Projects", names, default=names)
    st.success(f"Loaded {len(files)} file(s): " + ", ".join(f[0] for f in files))
sel = [r for r in reports if r.project in chosen] or reports
open_items = items[~items["status"].str.lower().isin(["closed", "done", "mitigated", "resolved"])] if len(items) else items

st.markdown(f'<div class="hero"><h1>📊 TPM Status Dashboard</h1><div class="s">As of {as_of:%A %d %B %Y} · {len(reports)} projects · '
            f'{len(items)} risks and issues · source: {", ".join(f[0] for f in files)}</div></div>', unsafe_allow_html=True)
for n in tracker.get("notes", []):
    st.warning(n)

tab_over, tab_rep, tab_reg, tab_data = st.tabs(["Overview", "Client reports", "Risk register", "Data loaded"])

# ------------------------------------------------------------------ overview
with tab_over:
    def stats(r):
        op = [i for i in r.all_items if i["status"].lower() not in ("closed", "done", "mitigated", "resolved", "complete")]
        return (sum(i["rating"] == "High" for i in op), sum(bool(i["due"] and i["due"] < as_of) for i in op))
    cnt = {k: sum(r.health == k for r in sel) for k in RAG}
    kp = [("Red projects", cnt["Red"], RAG["Red"]), ("Amber projects", cnt["Amber"], RAG["Amber"]), ("Green projects", cnt["Green"], RAG["Green"]),
          ("Open high risks", sum(stats(r)[0] for r in sel), "#0b2545"), ("Overdue actions", sum(stats(r)[1] for r in sel), "#c62828"),
          ("Decisions needed", sum(len(r.decisions) for r in sel), "#0b2545")]
    for col, (l, n, c) in zip(st.columns(6), kp):
        col.markdown(f'<div class="kpi"><div class="n" style="color:{c}">{n}</div><div class="l">{l}</div></div>', unsafe_allow_html=True)
    st.write("")
    for col, r in zip(st.columns(len(sel)), sel):
        m = r.milestones[0] if r.milestones else None
        days = (m["date"] - as_of).days if m and m["date"] else None
        when = "" if days is None else (f" · in {days} days" if days >= 0 else f" · {abs(days)} days overdue")
        col.markdown(f"""<div class="pcard"><span class="badge" style="background:{RAG.get(r.health, '#546e7a')}">{r.health.upper()}</span>
<h3>{r.project}</h3><div class="c">{r.client}</div>
<div class="row"><span class="k">Phase</span> · {r.phase or '–'}</div>
<div class="row"><span class="k">Next milestone</span> · {m['milestone'] if m else '–'}{when}</div>
<div class="row"><span class="k">Main issue</span> · {r.main_issue or '–'}</div>
<div class="row"><span class="k">Open high risks</span> · {stats(r)[0]} &nbsp; <span class="k">Overdue</span> · {stats(r)[1]}</div></div>""", unsafe_allow_html=True)
    st.write("")
    sel_names = [r.project for r in sel]
    oi = open_items[open_items["project"].apply(lambda p: any(p.lower() in n.lower() or n.lower() in p.lower() for n in sel_names))] if len(open_items) else open_items
    c1, c2 = st.columns(2)
    if len(oi):
        grid = [[0] * 3 for _ in range(3)]
        for _, row in oi.iterrows():
            if row["probability"] in LEVEL and row["impact"] in LEVEL:
                grid[LEVEL[row["probability"]] - 1][LEVEL[row["impact"]] - 1] += 1
        fig = go.Figure(go.Heatmap(z=[[(i + 1) * (j + 1) for j in range(3)] for i in range(3)], x=["Low", "Medium", "High"], y=["Low", "Medium", "High"],
                                   colorscale=[[0, "#cfe8d5"], [0.35, "#ffe7a8"], [1, "#f5b5b0"]], showscale=False, hoverinfo="skip",
                                   text=[[str(v) if v else "" for v in row] for row in grid], texttemplate="<b>%{text}</b>", textfont={"size": 22}))
        fig.update_layout(title="Risk heat map (open items)", xaxis_title="Impact", yaxis_title="Probability", height=330, margin=dict(l=10, r=10, t=45, b=10))
        stretch(c1.plotly_chart, fig)
        owners = oi.assign(owner=oi["owner"].replace("", "Unassigned")).groupby(["owner", "rating"]).size().reset_index(name="items")
        fig2 = px.bar(owners, y="owner", x="items", color="rating", orientation="h", title="Open items by owner",
                      color_discrete_map={"High": RAG["Red"], "Medium": RAG["Amber"], "Low": RAG["Green"], "Unrated": "#90a4ae"})
        fig2.update_layout(height=330, margin=dict(l=10, r=10, t=45, b=10), yaxis_title=None, legend_title=None)
        stretch(c2.plotly_chart, fig2)
        tl = oi.dropna(subset=["due"]).copy()
        if len(tl):
            tl["due"] = pd.to_datetime(tl["due"])
            fig3 = px.scatter(tl, x="due", y="project", color="rating", hover_data=["id", "owner", "status"], hover_name="description",
                              color_discrete_map={"High": RAG["Red"], "Medium": RAG["Amber"], "Low": RAG["Green"], "Unrated": "#90a4ae"}, title="Action due dates")
            fig3.update_traces(marker=dict(size=15, line=dict(width=1, color="white")))
            fig3.add_vline(x=pd.Timestamp(as_of).timestamp() * 1000, line_dash="dash", line_color="#0b2545", annotation_text="Report date")
            fig3.update_layout(height=300, margin=dict(l=10, r=10, t=45, b=10), yaxis_title=None, xaxis_title=None, legend_title=None)
            stretch(st.plotly_chart, fig3)
    else:
        st.info("No open risks or issues for the selected projects.")

# ------------------------------------------------------------------ client reports
with tab_rep:
    st.caption("Client-ready, one per client. Items marked Internal in the tracker and other projects' items are never shown here. Download, then print to PDF or send as HTML.")
    zip_buf = io.BytesIO()
    pages = {}
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for r in sel:
            slug = "".join(ch if ch.isalnum() else "_" for ch in r.client).strip("_")
            pages[r.client] = (slug, render_html(r), render_markdown(r), r)
            z.writestr(f"{slug}_status_{as_of:%Y-%m-%d}.html", pages[r.client][1])
            z.writestr(f"{slug}_status_{as_of:%Y-%m-%d}.md", pages[r.client][2])
    rtabs = st.tabs(list(pages))
    for t, (client, (slug, page, md, r)) in zip(rtabs, pages.items()):
        with t:
            for w in r.warnings:
                st.warning(w)
            embed_html(page, 1250)
            a, b = st.columns(2)
            a.download_button("⬇ Download report (HTML)", page, file_name=f"{slug}_status_{as_of:%Y-%m-%d}.html", mime="text/html", key=f"h_{slug}")
            b.download_button("⬇ Download Markdown", md, file_name=f"{slug}_status_{as_of:%Y-%m-%d}.md", mime="text/markdown", key=f"m_{slug}")
    st.download_button("⬇ Download all client reports (ZIP)", zip_buf.getvalue(), file_name=f"status_reports_{as_of:%Y-%m-%d}.zip", mime="application/zip")

# ------------------------------------------------------------------ register
with tab_reg:
    if len(items):
        f1, f2, f3 = st.columns(3)
        rf = f1.multiselect("Rating", ["High", "Medium", "Low", "Unrated"], default=["High", "Medium", "Low", "Unrated"])
        sf = f2.multiselect("Status", sorted(items["status"].unique()), default=sorted(items["status"].unique()))
        of = f3.multiselect("Owner", sorted(items["owner"].unique()), default=sorted(items["owner"].unique()))
        v = items[items["rating"].isin(rf) & items["status"].isin(sf) & items["owner"].isin(of)].copy()
        v["overdue"] = v["due"].apply(lambda d: bool(d and d < as_of)) & ~v["status"].str.lower().isin(["closed", "done", "mitigated"])
        show = v[["id", "project", "type", "description", "probability", "impact", "rating", "owner", "action", "due", "status", "visibility", "overdue"]]
        colours = {"High": "background-color:#f8b4b4;color:#8b0000;font-weight:600", "Medium": "background-color:#ffe08a;color:#7a4a00;font-weight:600",
                   "Low": "background-color:#b7e4c7;color:#14532d;font-weight:600"}
        stretch(st.dataframe, show.style.map(lambda x: colours.get(x, ""), subset=["rating"]), hide_index=True, height=480)
        st.download_button("⬇ Download register (CSV)", show.to_csv(index=False), file_name=f"risk_register_{as_of:%Y-%m-%d}.csv", mime="text/csv")
    else:
        st.info("No risks or issues in the tracker.")

# ------------------------------------------------------------------ data loaded
with tab_data:
    st.subheader("What the app read from your upload")
    stretch(st.dataframe, pd.DataFrame(tracker["sources"]), hide_index=True)
    if tracker.get("ignored"):
        st.info("Not used (no Summary, Risks or Milestones table found): " + ", ".join(tracker["ignored"]))
    st.markdown("**Summary rows**")
    stretch(st.dataframe, tracker["summary"].dropna(axis=1, how="all"), hide_index=True)
    st.markdown("**Risks and issues rows**")
    stretch(st.dataframe, tracker["risks"].dropna(axis=1, how="all"), hide_index=True)
