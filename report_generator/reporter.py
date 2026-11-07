"""Turn a TPM project tracker (Excel or CSV) into one weekly status report per client.

No AI/LLM calls: every line in a report comes from a cell in the tracker or from a
simple, visible rule (see README, "How the rules work").
"""
from __future__ import annotations

import html
import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import pandas as pd

SCORE = {"low": 1, "medium": 2, "high": 3}
CLOSED = {"closed", "done", "mitigated", "resolved", "complete", "completed"}

# column name (lower-case, as found in the file) contains ...  ->  internal name
SUMMARY_ALIASES = [
    ("client-facing", "client_update"), ("client update", "client_update"),
    ("rag", "status"), ("status", "status"), ("project", "project"), ("client", "client"),
    ("current phase", "phase"), ("phase", "phase"),
    ("next milestone", "next_milestone"), ("milestone date", "milestone_date"),
    ("main issue", "main_issue"), ("tpm action", "tpm_action"),
    ("completed", "completed"), ("decision", "decisions"),
]
RISK_ALIASES = [
    ("id", "id"), ("project", "project"), ("type", "type"), ("description", "description"),
    ("probability", "probability"), ("impact", "impact"), ("rating", "rating"),
    ("owner", "owner"), ("mitigation", "action"), ("action", "action"),
    ("due", "due"), ("visib", "visibility"), ("status", "status"),
]
MILESTONE_ALIASES = [("project", "project"), ("milestone", "milestone"), ("date", "date"), ("status", "status")]


class TrackerError(Exception):
    """A problem with the uploaded file that the user can fix. Message is shown as-is."""


@dataclass
class ClientReport:
    client: str
    project: str
    as_of: date
    health: str                      # Red / Amber / Green
    health_note: str
    phase: str
    progress: str
    completed: list[str] = field(default_factory=list)
    milestones: list[dict] = field(default_factory=list)
    risks: list[dict] = field(default_factory=list)
    overdue: list[dict] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    decisions: list[str] = field(default_factory=list)
    next_week: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    main_issue: str = ""
    all_items: list[dict] = field(default_factory=list)


# ----------------------------------------------------------------- loading
def _clean(v) -> str:
    return "" if pd.isna(v) else str(v).strip()


def _find_header(raw: pd.DataFrame, must_contain: str) -> int:
    for i in range(min(len(raw), 15)):
        cells = [str(c).strip().lower() for c in raw.iloc[i].tolist() if not pd.isna(c)]
        if any(must_contain in c for c in cells) and len(cells) >= 3:
            return i
    raise TrackerError(f"Could not find a header row containing '{must_contain}' in the first 15 rows.")


def _frame(raw: pd.DataFrame, aliases, key: str) -> pd.DataFrame:
    h = _find_header(raw, key)
    df = raw.iloc[h + 1:].copy()
    names = [str(c).strip().lower() if not pd.isna(c) else "" for c in raw.iloc[h].tolist()]
    mapped, used = [], set()
    for n in names:
        target = next((t for k, t in aliases if k in n and t not in used), None)
        if target:
            used.add(target)
        mapped.append(target or f"_extra_{len(mapped)}")
    df.columns = mapped
    # The table ends at the first fully blank row. Anything below it (legend, notes) is not data.
    blank = df.isna().all(axis=1).tolist()
    if True in blank:
        df = df.iloc[:blank.index(True)]
    return df.reset_index(drop=True)


def _read_any(name: str, data: bytes) -> dict[str, pd.DataFrame]:
    lower = name.lower()
    try:
        if lower.endswith((".xlsx", ".xlsm")):
            return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="openpyxl")
        if lower.endswith(".csv"):
            return {name: pd.read_csv(io.BytesIO(data), header=None, dtype=object)}
    except Exception as exc:  # unreadable / corrupt file
        raise TrackerError(f"'{name}' could not be read ({type(exc).__name__}). Is it a valid .xlsx or .csv file?") from exc
    raise TrackerError(f"'{name}' is not supported. Upload an .xlsx or .csv file.")


def load_tracker(files: list[tuple[str, bytes]]) -> dict:
    """files = [(file_name, file_bytes), ...]. Returns dict with summary, risks, milestones, raw_title."""
    if not files:
        raise TrackerError("Upload a tracker first.")
    sheets: dict[str, pd.DataFrame] = {}
    for name, data in files:
        if not data:
            raise TrackerError(f"'{name}' is empty.")
        for sname, df in _read_any(name, data).items():
            sheets[f"{name}::{sname}"] = df

    summaries, risk_frames, milestone_frames, sources, ignored = [], [], [], [], []
    title_text = ""
    for key, raw in sheets.items():
        fname, sname = key.split("::", 1)
        if raw.empty:
            ignored.append(f"{fname} / {sname}")
            continue
        flat = " ".join(str(x) for x in raw.head(15).fillna("").values.ravel()).lower()
        try:
            if "rag" in flat and "next milestone" in flat:
                df, kind = _frame(raw, SUMMARY_ALIASES, "project"), "Summary"
                summaries.append(df)
            elif "probability" in flat and "impact" in flat:
                df, kind = _frame(raw, RISK_ALIASES, "description"), "Risks and issues"
                risk_frames.append(df)
            elif "milestone" in flat and "date" in flat and "next milestone" not in flat:
                df, kind = _frame(raw, MILESTONE_ALIASES, "milestone"), "Milestones"
                milestone_frames.append(df)
            else:
                ignored.append(f"{fname} / {sname}")
                continue
        except TrackerError:
            ignored.append(f"{fname} / {sname}")
            continue
        title_text += " " + " ".join(str(x) for x in raw.head(3).fillna("").values.ravel())
        sources.append({"File": fname, "Sheet": sname, "Read as": kind, "Rows": len(df)})

    if not summaries:
        raise TrackerError("No Summary sheet found. It needs columns such as Project, Client, Status (RAG), Next milestone, Milestone date.")
    if not risk_frames:
        raise TrackerError("No 'Risks and Issues' sheet found. It needs columns such as Description, Probability, Impact, Owner, Due date, Status.")
    summary = pd.concat(summaries, ignore_index=True)
    risks = pd.concat(risk_frames, ignore_index=True)
    milestones = pd.concat(milestone_frames, ignore_index=True) if milestone_frames else None
    notes = []
    # the same project in two files: the file uploaded last wins (and we say so)
    dup = summary["project"].astype(str).str.strip().str.lower().duplicated(keep="last")
    if dup.any():
        notes.append("Project listed in more than one file (" + ", ".join(summary.loc[dup, "project"].astype(str)) + "): the later upload was used.")
        summary = summary[~dup].reset_index(drop=True)
    if "id" in risks.columns:
        key_ = risks["id"].astype(str).str.strip()
        d2 = key_.duplicated(keep="last") & key_.ne("nan") & key_.ne("")
        if d2.any():
            notes.append(f"{int(d2.sum())} risk/issue row(s) with the same ID in several files: the later upload was used.")
            risks = risks[~d2].reset_index(drop=True)
    for need in ("project", "client", "status"):
        if need not in summary.columns:
            raise TrackerError(f"The Summary sheet is missing the '{need}' column.")
    for need in ("project", "description", "probability", "impact"):
        if need not in risks.columns:
            raise TrackerError(f"The Risks and Issues sheet is missing the '{need}' column.")
    return {"summary": summary, "risks": risks, "milestones": milestones, "title_text": title_text,
            "ignored": ignored, "sources": sources, "notes": notes}


def detect_as_of(tracker: dict) -> date | None:
    m = re.search(r"as of\s+(?:\w{3}\s+)?(\d{1,2}\s+\w{3,9}\s+\d{4})", tracker.get("title_text", ""), re.I)
    if m:
        for fmt in ("%d %b %Y", "%d %B %Y"):
            try:
                return datetime.strptime(m.group(1), fmt).date()
            except ValueError:
                pass
    return None


# ----------------------------------------------------------------- logic
def _to_date(v) -> date | None:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    text = str(v).strip()
    iso = bool(re.match(r"\d{4}-\d{2}-\d{2}", text))      # ISO text is year-first; anything else is read day-first (India)
    d = pd.to_datetime(text, errors="coerce", dayfirst=not iso)
    return None if pd.isna(d) else d.date()


def _rating(row) -> str:
    given = _clean(row.get("rating", "")).lower()
    if given in SCORE:
        return given.capitalize()
    p, i = SCORE.get(_clean(row.get("probability")).lower()), SCORE.get(_clean(row.get("impact")).lower())
    if not p or not i:
        return "Unrated"
    s = p * i
    return "High" if s >= 6 else "Medium" if s >= 3 else "Low"


def _same_project(a: str, b: str) -> bool:
    a, b = a.lower().split("(")[0].strip(), b.lower().split("(")[0].strip()
    return bool(a) and (a == b or a in b or b in a)


def fmt(d: date | None) -> str:
    return d.strftime("%a %d %b") if d else "date not set"


def normalized_items(tracker: dict) -> list[dict]:
    items = []
    for _, r in tracker["risks"].iterrows():
        d = r.to_dict()
        items.append({
            "id": _clean(d.get("id")), "project": _clean(d.get("project")), "type": _clean(d.get("type")) or "Risk",
            "description": _clean(d.get("description")), "probability": _clean(d.get("probability")).capitalize(),
            "impact": _clean(d.get("impact")).capitalize(), "rating": _rating(d), "owner": _clean(d.get("owner")),
            "action": _clean(d.get("action")), "due": _to_date(d.get("due")), "status": _clean(d.get("status")),
            "visibility": _clean(d.get("visibility")) or "Client",
        })
    return items


def build_reports(tracker: dict, as_of: date) -> list[ClientReport]:
    summary, ms = tracker["summary"], tracker["milestones"]
    items = normalized_items(tracker)
    reports = []
    for _, s in summary.iterrows():
        project, client = _clean(s.get("project")), _clean(s.get("client"))
        if not project:
            continue
        mine = [i for i in items if _same_project(project, i["project"])]
        mine_all = mine
        mine = [i for i in mine_all if i["visibility"].lower() != "internal"]      # client reports never show Internal items
        open_ = [i for i in mine if i["status"].lower() not in CLOSED]
        warnings = []

        major = sorted([i for i in open_ if i["rating"] == "High"], key=lambda i: (i["due"] or date.max))
        overdue = sorted([i for i in open_ if i["due"] and i["due"] < as_of], key=lambda i: i["due"])
        for i in mine:
            if i["due"] is None and i["status"].lower() not in CLOSED:
                warnings.append(f"{i['id'] or i['description'][:30]}: due date missing or not readable.")

        reported = _clean(s.get("status")).capitalize()
        if reported not in ("Red", "Amber", "Green"):
            warnings.append(f"Status '{reported}' is not Red/Amber/Green; health was calculated from risks.")
            reported = ""
        calc = "Red" if any(i in overdue and i["rating"] == "High" for i in open_) else \
               "Amber" if major or overdue else "Green"
        health = reported or calc
        order = {"Green": 0, "Amber": 1, "Red": 2}
        note = "Status as set in the tracker."
        if reported and order[calc] > order[reported]:
            note = f"Tracker says {reported}, but open items suggest {calc} (overdue or high-rated items). Review before sending."
            warnings.append(note)

        milestones = []
        nm, nd = _clean(s.get("next_milestone")), _to_date(s.get("milestone_date"))
        if nm:
            milestones.append({"milestone": nm, "date": nd, "overdue": bool(nd and nd < as_of)})
        if ms is not None:
            for _, m in ms.iterrows():
                if _same_project(project, _clean(m.get("project"))) and _clean(m.get("status")).lower() not in CLOSED:
                    md = _to_date(m.get("date"))
                    if md and md <= as_of + timedelta(days=14):
                        milestones.append({"milestone": _clean(m.get("milestone")), "date": md, "overdue": md < as_of})
        milestones.sort(key=lambda m: m["date"] or date.max)

        done = [i for i in mine if i["status"].lower() in CLOSED]
        completed = [t.strip().rstrip(".;") for t in re.split(r"(?<=[.;])\s+", _clean(s.get("completed"))) if t.strip()]
        completed += [f"{i['id']}: {i['description']} (closed)" for i in done]

        blockers = [f"{i['id']} {i['description']}" for i in open_ if i["type"].lower() == "issue" and i["rating"] == "High"]
        decisions = [t.strip().rstrip(".") for t in re.split(r"\.\s+(?=[A-Z])", _clean(s.get("decisions"))) if t.strip()]
        week_end = as_of + timedelta(days=7)
        nxt = sorted([i for i in open_ if i["due"] and as_of <= i["due"] <= week_end], key=lambda i: i["due"])

        progress = _clean(s.get("client_update"))
        if not progress:
            progress = (f"Phase: {_clean(s.get('phase'))}. " if _clean(s.get("phase")) else "") + (f"Main issue: {_clean(s.get('main_issue'))}" if _clean(s.get("main_issue")) else "")
            warnings.append("No 'Client-facing update' text in the tracker, so the internal 'Main issue' text is used in the client report. Check it before sending.")
        reports.append(ClientReport(
            client=client or project, project=project, as_of=as_of, health=health, health_note=note,
            phase=_clean(s.get("phase")),
            progress=progress,
            completed=completed, milestones=milestones, risks=major, overdue=overdue, blockers=blockers,
            decisions=decisions, next_week=nxt, warnings=warnings,
            main_issue=_clean(s.get("main_issue")), all_items=mine_all))
    if not reports:
        raise TrackerError("The Summary sheet has no project rows.")
    return reports


# ----------------------------------------------------------------- rendering
def render_markdown(r: ClientReport) -> str:
    L = [f"# Weekly Status Report – {r.client}", f"**Project:** {r.project}  |  **As of:** {r.as_of:%a %d %b %Y}",
         "", f"## Project health: {r.health.upper()}", "", "## Progress summary", r.progress or "No summary provided.", "",
         "## Completed"]
    L += [f"- {c}" for c in r.completed] or ["- Nothing recorded this week."]
    L += ["", "## Upcoming milestones"]
    L += [f"- **{fmt(m['date'])}** – {m['milestone']}" + (" ⚠ overdue" if m["overdue"] else "") for m in r.milestones] or ["- None in the next 14 days."]
    L += ["", "## Major risks and issues", "| Item | Rating | Owner | Action | Due |", "|---|---|---|---|---|"]
    L += [f"| {i['id']} {i['description']} | {i['rating']} | {i['owner']} | {i['action']} | {fmt(i['due'])} |" for i in r.risks] or ["| None | | | | |"]
    L += ["", "## Blockers and dependencies"]
    L += [f"- {b}" for b in r.blockers] or ["- No blocking issues."]
    if r.overdue:
        L += ["", "## Overdue actions"] + [f"- {i['id']} (due {fmt(i['due'])}, owner {i['owner']}): {i['action']}" for i in r.overdue]
    L += ["", "## Decisions required"] + ([f"- {d}" for d in r.decisions] or ["- None this week."])
    L += ["", "## Next week's action plan", "| Due | Action | Owner |", "|---|---|---|"]
    L += [f"| {fmt(i['due'])} | {i['id']} – {i['action']} | {i['owner']} |" for i in r.next_week] or ["| – | No actions due | |"]
    return "\n".join(L) + "\n"


def render_html(r: ClientReport) -> str:
    """A client-ready, print-friendly one-page report (open in a browser, Print > Save as PDF)."""
    e = html.escape
    colour = {"Red": "#c62828", "Amber": "#ef8f00", "Green": "#2e7d32"}.get(r.health, "#546e7a")
    soft = {"Red": "#fdecea", "Amber": "#fff4e0", "Green": "#e8f5e9"}.get(r.health, "#eceff1")
    rate = {"High": "#c62828", "Medium": "#ef8f00", "Low": "#2e7d32"}

    def card(title, inner):
        return f'<section class="card"><h2>{title}</h2>{inner}</section>'

    def ul(xs):
        return "<ul>" + "".join(f"<li>{x}</li>" for x in xs) + "</ul>" if xs else '<p class="muted">None.</p>'

    def table(head, rows):
        if not rows:
            return '<p class="muted">None.</p>'
        return ("<table><tr>" + "".join(f"<th>{h}</th>" for h in head) + "</tr>"
                + "".join("<tr>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>" for row in rows) + "</table>")

    def pill(rating):
        return f'<span class="pill" style="background:{rate.get(rating, "#546e7a")}">{e(rating)}</span>'

    kpis = [("Major risks", len(r.risks)), ("Overdue actions", len(r.overdue)),
            ("Decisions needed", len(r.decisions)), ("Actions next 7 days", len(r.next_week))]
    kpi_html = "".join(f'<div class="kpi"><div class="n">{n}</div><div class="l">{l}</div></div>' for l, n in kpis)
    risks_t = table(["Item", "Rating", "Owner", "Action", "Due"],
                    [[f"<b>{e(i['id'])}</b> {e(i['description'])}", pill(i["rating"]), e(i["owner"]), e(i["action"]), fmt(i["due"])] for i in r.risks])
    ms = ul([f"<b>{fmt(m['date'])}</b> – {e(m['milestone'])}" + (' <span class="late">overdue</span>' if m["overdue"] else "") for m in r.milestones])
    nxt = table(["Due", "Action", "Owner"], [[fmt(i["due"]), f"<b>{e(i['id'])}</b> {e(i['action'])}", e(i["owner"])] for i in r.next_week])
    overdue = card("Overdue actions", ul([e(f"{i['id']} – due {fmt(i['due'])}, owner {i['owner']}: {i['action']}") for i in r.overdue])) if r.overdue else ""
    body = f"""
<header style="border-left:8px solid {colour};background:{soft}">
  <div><div class="eyebrow">Weekly status report</div><h1>{e(r.client)}</h1>
  <div class="muted">Project: <b>{e(r.project)}</b> &nbsp;|&nbsp; Week of {r.as_of:%d %b %Y} &nbsp;|&nbsp; Phase: {e(r.phase) or '–'}</div></div>
  <div class="health" style="background:{colour}">{r.health.upper()}</div>
</header>
<div class="kpis">{kpi_html}</div>
{card("Progress summary", f"<p>{e(r.progress) or 'No summary provided.'}</p>")}
<div class="two">{card("Completed", ul([e(c) for c in r.completed]))}{card("Upcoming milestones", ms)}</div>
{card("Major risks and issues", risks_t)}
<div class="two">{card("Blockers and dependencies", ul([e(b) for b in r.blockers]))}{card("Decisions required from you", ul([e(d) for d in r.decisions]))}</div>
{overdue}
{card("Next week's action plan", nxt)}
<footer>Prepared by the Technical Program Manager, PanScience Innovations. Source: project tracker as of {r.as_of:%d %b %Y}.</footer>"""
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>Status – {e(r.client)}</title><style>
*{{box-sizing:border-box}}body{{font-family:"Segoe UI",Arial,sans-serif;background:#f4f6f9;margin:0;padding:18px;color:#1f2933;font-size:13.5px}}
.wrap{{max-width:980px;margin:auto}}header{{display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-radius:10px;margin-bottom:12px}}
.eyebrow{{text-transform:uppercase;letter-spacing:.08em;font-size:11px;color:#52606d}}h1{{margin:2px 0 4px;font-size:26px;color:#0b2545}}
.health{{color:#fff;font-weight:700;font-size:20px;padding:10px 22px;border-radius:30px;letter-spacing:.05em}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px}}.kpi{{background:#fff;border-radius:10px;padding:10px 14px;box-shadow:0 1px 3px rgba(0,0,0,.08)}}
.kpi .n{{font-size:26px;font-weight:700;color:#0b2545}}.kpi .l{{color:#52606d;font-size:12px}}
.card{{background:#fff;border-radius:10px;padding:12px 16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.08)}}
.card h2{{margin:0 0 6px;font-size:14px;color:#0b2545;text-transform:uppercase;letter-spacing:.05em}}.two{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
table{{border-collapse:collapse;width:100%}}th{{text-align:left;font-size:11.5px;color:#52606d;border-bottom:2px solid #e4e7eb;padding:5px 6px}}
td{{border-bottom:1px solid #eef0f3;padding:6px;vertical-align:top}}ul{{margin:4px 0 4px 18px;padding:0}}li{{margin:3px 0}}
.pill{{color:#fff;border-radius:10px;padding:1px 9px;font-size:11.5px;font-weight:600}}.muted{{color:#7b8794}}.late{{color:#c62828;font-weight:600}}
footer{{color:#7b8794;font-size:11.5px;text-align:center;margin-top:8px}}
@media print{{body{{background:#fff;padding:0}}.card,.kpi{{box-shadow:none;border:1px solid #e4e7eb}}}}
</style></head><body><div class="wrap">{body}</div></body></html>"""
