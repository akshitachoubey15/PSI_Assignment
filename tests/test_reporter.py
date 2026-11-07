import io
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from report_generator import TrackerError, build_reports, detect_as_of, load_tracker, render_html, render_markdown

SAMPLE = Path(__file__).parent.parent / "sample_data" / "sample_tracker.xlsx"


def sample(as_of=date(2026, 11, 9)):
    t = load_tracker([(SAMPLE.name, SAMPLE.read_bytes())])
    return t, build_reports(t, as_of)


def test_one_report_per_project_and_no_legend_rows():
    # Regression: the "Status key" legend under the Summary table was once read as 4 extra projects.
    _, reports = sample()
    assert [r.project for r in reports] == ["Kavach", "Ledger"]
    assert [r.client for r in reports] == ["Nirmal Finance", "Sagar Publications"]


def test_health_and_milestone():
    _, (k, l) = sample()
    assert k.health == "Red" and l.health == "Amber"
    assert k.milestones[0]["date"] == date(2026, 11, 12)


def test_reports_do_not_leak_other_clients_items():
    _, (k, l) = sample()
    assert all(i["id"].startswith("K-") for i in k.risks + k.next_week)
    assert all(i["id"].startswith("L-") for i in l.risks + l.next_week)   # "Both" items (X-01) stay internal


def test_overdue_uses_report_date():
    _, (k, _) = sample(date(2026, 11, 12))
    assert {"K-06", "K-02", "K-03", "K-04"} <= {i["id"] for i in k.overdue}
    _, (k0, _) = sample(date(2026, 11, 9))
    assert k0.overdue == []


def test_detect_as_of_from_title():
    t, _ = sample()
    assert detect_as_of(t) == date(2026, 11, 9)


def test_markdown_and_html_render():
    _, (k, _) = sample()
    md, page = render_markdown(k), render_html(k)
    assert "Project health: RED" in md and "Decisions required" in md
    assert "<html" in page and "Nirmal Finance" in page


def test_csv_upload_two_files():
    sheets = pd.read_excel(SAMPLE, sheet_name=None, header=None)
    s_csv = sheets["Summary"].to_csv(index=False, header=False).encode()
    r_csv = sheets["Risks and Issues"].to_csv(index=False, header=False).encode()
    t2 = load_tracker([("summary.csv", s_csv), ("risks.csv", r_csv)])
    assert [r.project for r in build_reports(t2, date(2026, 11, 9))] == ["Kavach", "Ledger"]


def test_missing_sheet_gives_friendly_error():
    buf = io.BytesIO()
    pd.DataFrame({"a": [1]}).to_excel(buf, index=False)
    with pytest.raises(TrackerError, match="No Summary sheet"):
        load_tracker([("bad.xlsx", buf.getvalue())])


def test_bad_file_type_and_empty_file():
    with pytest.raises(TrackerError, match="not supported"):
        load_tracker([("notes.docx", b"x")])
    with pytest.raises(TrackerError, match="empty"):
        load_tracker([("x.csv", b"")])
    with pytest.raises(TrackerError):
        load_tracker([])


def test_unreadable_due_date_becomes_warning_not_crash():
    t, _ = sample()
    t["risks"].loc[0, "due"] = "soon"
    k = build_reports(t, date(2026, 11, 9))[0]
    assert any("due date" in w for w in k.warnings)


def test_tracker_status_worse_than_reported_is_flagged():
    t, _ = sample(date(2026, 11, 13))
    t["summary"].loc[1, "status"] = "Green"      # marked Green although a High item (L-04, due 13 Nov) is overdue on 20 Nov
    l = build_reports(t, date(2026, 11, 20))[1]
    assert l.health == "Green" and any("suggest" in w for w in l.warnings)
