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
    ("due", "due"), ("status", "status"),
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

    summary = risks = milestones = None
    title_text = ""
    ignored = []
    for key, raw in sheets.items():
        if raw.empty:
            ignored.append(key.split("::")[0])
            continue
        flat = " ".join(str(x) for x in raw.head(15).fillna("").values.ravel()).lower()
        title_text += " " + " ".join(str(x) for x in raw.head(3).fillna("").values.ravel())
        try:
            if "rag" in flat and "next milestone" in flat:
                summary = _frame(raw, SUMMARY_ALIASES, "project")
            elif "probability" in flat and "impact" in flat:
                risks = _frame(raw, RISK_ALIASES, "description")
            elif "milestone" in flat and "date" in flat and "next milestone" not in flat:
                milestones = _frame(raw, MILESTONE_ALIASES, "milestone")
            else:
                ignored.append(key)
        except TrackerError:
            ignored.append(key)
    if summary is None:
        raise TrackerError("No Summary sheet found. It needs columns such as Project, Client, Status (RAG), Next milestone, Milestone date.")
    for need in ("project", "client", "status"):
        if need not in summary.columns:
            raise TrackerError(f"The Summary sheet is missing the '{need}' column.")
    if risks is None:
        raise TrackerError("No 'Risks and Issues' sheet found. It needs columns such as Description, Probability, Impact, Owner, Due date, Status.")
    for need in ("project", "description", "probability", "impact"):
        if need not in risks.columns:
            raise TrackerError(f"The Risks and Issues sheet is missing the '{need}' column.")
    return {"summary": summary, "risks": risks, "milestones": milestones, "title_text": title_text, "ignored": ignored}


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


def build_reports(tracker: dict, as_of: date) -> list[ClientReport]:
    summary, risks, ms = tracker["summary"], tracker["risks"], tracker["milestones"]
    items = []
    for _, r in risks.iterrows():
        d = r.to_dict()
        items.append({
            "id": _clean(d.get("id")), "project": _clean(d.get("project")), "type": _clean(d.get("type")) or "Risk",
            "description": _clean(d.get("description")), "rating": _rating(d), "owner": _clean(d.get("owner")),
            "action": _clean(d.get("action")), "due": _to_date(d.get("due")), "status": _clean(d.get("status")),
        })
    reports = []
    for _, s in summary.iterrows():
        project, client = _clean(s.get("project")), _clean(s.get("client"))
        if not project:
            continue
        mine = [i for i in items if _same_project(project, i["project"])]
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

        reports.append(ClientReport(
            client=client or project, project=project, as_of=as_of, health=health, health_note=note,
            phase=_clean(s.get("phase")),
            progress=(f"Phase: {_clean(s.get('phase'))}. " if _clean(s.get("phase")) else "")
                     + (f"Main issue: {_clean(s.get('main_issue'))}" if _clean(s.get("main_issue")) else ""),
            completed=completed, milestones=milestones, risks=major, overdue=overdue, blockers=blockers,
            decisions=decisions, next_week=nxt, warnings=warnings))
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
    e = html.escape
    colour = {"Red": "#b00020", "Amber": "#b26a00", "Green": "#1b7a2f"}.get(r.health, "#444")

    def ul(xs): return "<ul>" + "".join(f"<li>{x}</li>" for x in xs) + "</ul>" if xs else "<p>None.</p>"

    def table(head, rows):
        if not rows:
            return "<p>None.</p>"
        return ("<table><tr>" + "".join(f"<th>{h}</th>" for h in head) + "</tr>"
                + "".join("<tr>" + "".join(f"<td>{e(str(c))}</td>" for c in row) + "</tr>" for row in rows) + "</table>")
    body = f"""<h1>Weekly Status Report – {e(r.client)}</h1>
<p>Project: <b>{e(r.project)}</b> | As of {r.as_of:%a %d %b %Y}</p>
<p>Project health: <span class="badge" style="background:{colour}">{r.health.upper()}</span></p>
<h2>Progress summary</h2><p>{e(r.progress) or 'No summary provided.'}</p>
<h2>Completed</h2>{ul([e(c) for c in r.completed])}
<h2>Upcoming milestones</h2>{ul([f"<b>{fmt(m['date'])}</b> – {e(m['milestone'])}" + (' ⚠ overdue' if m['overdue'] else '') for m in r.milestones])}
<h2>Major risks and issues</h2>{table(['Item','Rating','Owner','Action','Due'], [[f"{i['id']} {i['description']}", i['rating'], i['owner'], i['action'], fmt(i['due'])] for i in r.risks])}
<h2>Blockers and dependencies</h2>{ul([e(b) for b in r.blockers])}
{('<h2>Overdue actions</h2>' + ul([e(f"{i['id']} (due {fmt(i['due'])}, owner {i['owner']}): {i['action']}") for i in r.overdue])) if r.overdue else ''}
<h2>Decisions required</h2>{ul([e(d) for d in r.decisions])}
<h2>Next week's action plan</h2>{table(['Due','Action','Owner'], [[fmt(i['due']), f"{i['id']} – {i['action']}", i['owner']] for i in r.next_week])}"""
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>Status – {e(r.client)}</title><style>
body{{font-family:Arial,sans-serif;max-width:900px;margin:24px auto;color:#1a1a1a;font-size:14px}}
h1{{color:#0b2545}}h2{{color:#0b2545;border-bottom:1px solid #0b2545;padding-bottom:2px;margin-top:20px}}
table{{border-collapse:collapse;width:100%}}th{{background:#0b2545;color:#fff;text-align:left;padding:4px 6px}}
td{{border:1px solid #c9cfd8;padding:4px 6px;vertical-align:top}}.badge{{color:#fff;padding:2px 10px;border-radius:10px;font-weight:bold}}
</style></head><body>{body}</body></html>"""
