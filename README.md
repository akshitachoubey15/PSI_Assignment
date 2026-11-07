# TPM Status Dashboard and Report Generator

A Streamlit dashboard for Technical Program Managers. Upload a project tracker (Excel or CSV) and get a portfolio overview, a risk register and **one client-ready weekly report per client**. It was built for the PSI take-home, where a Friday client report was missed.

![screenshot](docs/screenshot.png)

**Status (honest):** built with Claude; tested on a Linux sandbox (Python 3.13). 15 unit tests pass. The app was driven in headless Chrome: upload, then upload an *edited* file and see the dashboard change, bad file, ZIP download. It is **not deployed** and **not yet pushed to GitHub** (this folder is the repository content).

## Run it

```bash
python -m venv .venv && source .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py          # opens http://localhost:8501
python -m pytest -q           # optional
```

1. Upload your `.xlsx` (or a Summary `.csv` and a Risks `.csv`) in the sidebar. The sample tracker is only for a first look and switches off when you upload.
2. The **Data loaded** tab shows exactly which sheets and rows were read, so you can see your upload took effect.
3. Optionally set **Report as of**; overdue items and "next 7 days" count from it.
4. Tabs: **Overview** (KPI cards, project cards, risk heat map, items by owner, due-date timeline), **Client reports** (polished report per client, download HTML/Markdown, or all as a ZIP), **Risk register** (filters, download CSV), **Data loaded**.

## What technology and APIs it uses

No external API, no API key, no internet call and no AI model inside the app. Your tracker stays on your machine.

| Part | Library | Used for |
|---|---|---|
| Interface | Streamlit | Pages, file upload, tabs, downloads |
| Data | pandas + openpyxl | Reading Excel/CSV, cleaning, rules |
| Charts | Plotly | Heat map, bar chart, timeline |
| Reports | plain HTML/CSS written in `reporter.py` | Client report (print to PDF from the browser) |
| Tests | pytest | 15 tests |

## How it was built (process)

1. Defined the input: the Task 5 tracker (Summary + Risks sheets).
2. Wrote `reporter.py` (no Streamlit inside): load files, find the header row, rename columns, apply the rules, render reports.
3. Wrote `app.py` on top: sidebar, tabs, charts.
4. Wrote tests first for each bug, ran the app in a real browser, and fixed what broke (see the errors below).

## Input format

The Task 5 workbook works as is. The tool finds the header row, so a title row above the table is fine.

| Sheet | Needed columns | Optional |
|---|---|---|
| Summary | Project, Client, Status (RAG), Next milestone, Milestone date | Current phase, Main issue, Completed so far, Decision required (client), **Client-facing update** |
| Risks and Issues | Project, Description, Probability, Impact | ID, Type, Risk rating, Owner, Mitigation / action, Due date, Status, **Visibility** (Client or Internal) |
| Milestones (extra sheet) | Project, Milestone, Date | Status |

## How the rules work

- **One report per Summary row.** Risks join by project name. Several files are merged; if the same project or risk ID is in two files, the later upload wins and the app says so.
- **Client-safe:** a client report shows only that project's items, never rows marked `Internal`, and uses the *Client-facing update* text. If that column is missing, the app warns you that internal text is being used.
- **Health:** the RAG in the tracker; the app warns if open items suggest a worse colour.
- **Rating:** from Probability x Impact (1-3 each; 6+ High, 3-4 Medium) when the cell is empty.
- **Major risks:** open High items. **Overdue:** open items due before the report date. **Next week:** open items due in 7 days.
- **Errors:** wrong type, empty file, missing sheet/column, unreadable dates and unused files are reported in plain words.

## Repository structure

```
tpm-status-report-generator/
├── app.py                       # Streamlit interface
├── report_generator/
│   ├── __init__.py
│   └── reporter.py              # load, rules, Markdown/HTML rendering
├── sample_data/sample_tracker.xlsx   # the Task 5 tracker (Kavach + Ledger)
├── tests/test_reporter.py       # 15 tests
├── docs/screenshot.png
├── requirements.txt
└── README.md
```

## Prompts used

Built with Claude (Claude Code environment). The prompt that defined the tool, from the assignment brief:

> **Prompt 1 – the build brief.** "Build a small tool that saves time for a TPM. It reads the tracker from Task 5 (CSV or Excel) and makes a weekly status report for each client: client name, project health, progress summary, completed activities, upcoming milestones, major risks, decisions required, next week's action plan. Highlight blockers, dependencies and actions. Python with Streamlit and Pandas. Easy interface, upload Excel/CSV, client-specific reports, download, sample data, error handling, README. Do not claim it is tested, pushed or deployed unless that happened."

Follow-up steps I asked for during the build: *"Run the app and check the output on the sample data"*, *"Write tests, including one for the bug you just found"*, *"Open the real app in a browser, upload a file, and try a bad file."*

## Errors the AI made, and how they were corrected

**1. Legend rows read as projects.** The first version made six reports instead of two: the colour key under the Summary table ("Status key: Red / Amber / Green") was read as four extra projects. I found it by printing the output, not by reading code. The AI had assumed the table runs to the end of the sheet and used `dropna(how="all")`, which removed the blank separator row. The table now ends at the first blank row, with a regression test.

**2. Internal text leaked into client reports.** The first dashboard put the tracker's internal notes (a team member's leave, the AI cost overrun) into the client report. A test that searched the client HTML for internal words caught it. Fix: a `Visibility` column, a `Client-facing update` column, and tests that no internal item or other client's name appears.

**3. A variable hid a function.** In the new dashboard I named a helper `show()` and later a DataFrame `show`, so the app crashed with "DataFrame object is not callable" the first time I opened it in the browser. Renamed the helper to `stretch()`.

**4. Smaller:** ISO dates were parsed day-first and Pandas warned; and a missing column header in the sample file (the new `Visibility` column) meant the data was ignored until a test failed.

## Limits

- Matching of Summary to Risks is by project name; keep names consistent.
- The report text is taken from your cells; the tool does not rewrite or summarise it, so it is only as good as the tracker.
- "Dependencies" are shown through blockers and decisions; the tool does not read text to find dependencies.
- The client report is HTML. Use your browser's Print > Save as PDF for a PDF copy.
