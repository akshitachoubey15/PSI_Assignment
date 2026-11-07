"""TPM Weekly Status Report Generator – Streamlit app.  Run: streamlit run app.py"""
import io
import zipfile
from datetime import date
from pathlib import Path

import streamlit as st

from report_generator import TrackerError, build_reports, detect_as_of, load_tracker, render_html, render_markdown

SAMPLE = Path(__file__).parent / "sample_data" / "sample_tracker.xlsx"
COLOURS = {"Red": "#b00020", "Amber": "#b26a00", "Green": "#1b7a2f"}

st.set_page_config(page_title="TPM Weekly Status Report Generator", page_icon="📋", layout="wide")
st.title("📋 TPM Weekly Status Report Generator")
st.caption("Upload your project tracker. Get one weekly status report per client, ready to download. No AI calls, no data leaves this machine.")

with st.sidebar:
    st.header("1. Tracker")
    uploaded = st.file_uploader("Excel (.xlsx) or CSV. For CSV, upload the Summary and the Risks files together.",
                                type=["xlsx", "csv"], accept_multiple_files=True)
    use_sample = st.checkbox("Use the sample tracker (Kavach + Ledger)", value=not uploaded)
    st.header("2. Report date")
    as_of_override = st.date_input("Report as of", value=None, help="Leave empty to use the date written in the tracker, or today.")
    st.caption("Overdue items and 'next week' are counted from this date.")

files = []
if uploaded:
    files = [(f.name, f.getvalue()) for f in uploaded]      # getvalue(): safe to read more than once
elif use_sample:
    files = [(SAMPLE.name, SAMPLE.read_bytes())]

if not files:
    st.info("Upload a tracker in the sidebar, or tick the sample tracker, to begin.")
    st.stop()

try:
    tracker = load_tracker(files)
    as_of = as_of_override or detect_as_of(tracker) or date.today()
    reports = build_reports(tracker, as_of)
except TrackerError as err:
    st.error(str(err))
    st.stop()
except Exception as err:                                      # last line of defence: never show a traceback to the user
    st.error(f"Something unexpected went wrong while reading the tracker ({type(err).__name__}). Check the file format against the sample.")
    st.stop()

if tracker.get("ignored"):
    st.info("Not used (no Summary, Risks or Milestones table found): " + ", ".join(tracker["ignored"]))
st.subheader(f"Portfolio view (internal) – as of {as_of:%a %d %b %Y}")
cols = st.columns(len(reports))
for c, r in zip(cols, reports):
    c.markdown(f"**{r.project}** – {r.client}  \n<span style='background:{COLOURS.get(r.health, '#444')};color:#fff;padding:2px 10px;border-radius:10px'>{r.health}</span>",
               unsafe_allow_html=True)
    c.caption(f"{len(r.risks)} major risks · {len(r.overdue)} overdue · {len(r.decisions)} decisions")

st.divider()
st.subheader("Client reports")
tabs = st.tabs([f"{r.client}" for r in reports])
zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
    for tab, r in zip(tabs, reports):
        md, page = render_markdown(r), render_html(r)
        slug = "".join(ch if ch.isalnum() else "_" for ch in r.client).strip("_")
        z.writestr(f"{slug}_status_{as_of:%Y-%m-%d}.md", md)
        z.writestr(f"{slug}_status_{as_of:%Y-%m-%d}.html", page)
        with tab:
            for w in r.warnings:
                st.warning(w)
            st.markdown(md)
            a, b = st.columns(2)
            a.download_button("Download Markdown", md, file_name=f"{slug}_status_{as_of:%Y-%m-%d}.md", mime="text/markdown", key=f"md_{slug}")
            b.download_button("Download HTML (print to PDF)", page, file_name=f"{slug}_status_{as_of:%Y-%m-%d}.html", mime="text/html", key=f"html_{slug}")
st.download_button("⬇ Download all reports (ZIP)", zip_buf.getvalue(), file_name=f"status_reports_{as_of:%Y-%m-%d}.zip", mime="application/zip")
