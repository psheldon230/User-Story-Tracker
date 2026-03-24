"""
End-to-end tests for all 3 Story Tracker tabs.

Tab 1 – Full Sync  : POST /sync  with jira_csv + excel_file
Tab 2 – VDI        : POST /parse-csv  with jira_csv  (server route kept for
                     backwards compat; client-side parse tested via unit test)
Tab 3 – Office     : POST /sync  with stories_json_text (pasted) + excel_file
                     POST /sync  with stories_json (file)  + excel_file

Sprint normalisation is tested independently.
"""
import io
import json
import re
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import openpyxl
import pytest

import app as app_module
from app import app, _normalise_sprint


# ── Helpers ───────────────────────────────────────────────────────────────────

def make_excel(rows) -> bytes:
    """Build a minimal .xlsx with given rows (list-of-lists). Returns bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_excel_bytes(data: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb.active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def make_csv(*rows) -> bytes:
    lines = [",".join(r) for r in rows]
    return "\n".join(lines).encode()


@pytest.fixture
def client():
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


# ── Sprint normalisation unit tests ───────────────────────────────────────────

class TestSprintNormalisation:
    cases = [
        ("Y26.SP04",     "Sprint 4"),
        ("Y26.SP12",     "Sprint 12"),
        ("SP05",         "Sprint 5"),
        ("SP1",          "Sprint 1"),
        ("Sprint 3",     "Sprint 3"),
        ("sprint03",     "Sprint 3"),
        ("sprint 10",    "Sprint 10"),
        ("Custom board", "Custom board"),   # unknown — pass through
        ("",             ""),
    ]

    @pytest.mark.parametrize("raw,expected", cases)
    def test_normalise(self, raw, expected):
        assert _normalise_sprint(raw) == expected


# ── Tab 1: Full Sync (/sync with jira_csv + excel_file) ──────────────────────

class TestFullSync:
    def test_missing_both_files_redirects(self, client):
        r = client.post("/sync")
        assert r.status_code == 303
        assert b"error" in r.headers["Location"].encode()

    def test_missing_excel_redirects(self, client):
        csv_bytes = make_csv(
            ["Issue key", "Sprint"],
            ["PROJ-1", "Y26.SP04"],
        )
        r = client.post("/sync", data={"jira_csv": (io.BytesIO(csv_bytes), "jira.csv")},
                        content_type="multipart/form-data")
        assert r.status_code == 303
        assert "error" in r.headers["Location"]

    def test_missing_csv_redirects(self, client):
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post("/sync", data={"excel_file": (io.BytesIO(xl), "tracker.xlsx")},
                        content_type="multipart/form-data")
        assert r.status_code == 303
        assert "error" in r.headers["Location"]

    def test_sync_adds_new_stories(self, client):
        csv_bytes = make_csv(
            ["Issue key", "Sprint"],
            ["PROJ-1", "Y26.SP04"],
            ["PROJ-2", "Y26.SP05"],
        )
        xl = make_excel([["Sprint", "Issue Key"], ["Sprint 4", "PROJ-1"]])
        r = client.post(
            "/sync",
            data={
                "jira_csv":    (io.BytesIO(csv_bytes), "jira.csv"),
                "excel_file":  (io.BytesIO(xl),        "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        rows = read_excel_bytes(r.data)
        keys = [str(row[1]) for row in rows[1:] if row[1]]
        assert "PROJ-1" in keys
        assert "PROJ-2" in keys

    def test_sprint_normalised_in_output(self, client):
        csv_bytes = make_csv(
            ["Issue key", "Sprint"],
            ["PROJ-3", "Y26.SP07"],
        )
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={
                "jira_csv":   (io.BytesIO(csv_bytes), "jira.csv"),
                "excel_file": (io.BytesIO(xl),        "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        rows = read_excel_bytes(r.data)
        sprints = [str(row[0]) for row in rows[1:] if row[0]]
        assert "Sprint 7" in sprints

    def test_bad_csv_no_key_column_redirects(self, client):
        csv_bytes = make_csv(["Summary", "Status"], ["A story", "Open"])
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={
                "jira_csv":   (io.BytesIO(csv_bytes), "jira.csv"),
                "excel_file": (io.BytesIO(xl),        "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 303
        assert "error" in r.headers["Location"]


# ── Tab 2: VDI – /parse-csv ───────────────────────────────────────────────────

class TestVdiParseCsv:
    def test_returns_json(self, client):
        csv_bytes = make_csv(
            ["Issue key", "Sprint"],
            ["PROJ-1", "Y26.SP04"],
            ["PROJ-2", "Y26.SP05"],
        )
        r = client.post(
            "/parse-csv",
            data={"jira_csv": (io.BytesIO(csv_bytes), "jira.csv")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        data = json.loads(r.data)
        assert isinstance(data, list)
        assert len(data) == 2

    def test_sprint_normalised_in_json(self, client):
        csv_bytes = make_csv(
            ["Issue key", "Sprint"],
            ["PROJ-1", "Y26.SP04"],
            ["PROJ-2", "SP12"],
        )
        r = client.post(
            "/parse-csv",
            data={"jira_csv": (io.BytesIO(csv_bytes), "jira.csv")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        data = json.loads(r.data)
        sprints = {item["key"]: item["sprint"] for item in data}
        assert sprints["PROJ-1"] == "Sprint 4"
        assert sprints["PROJ-2"] == "Sprint 12"

    def test_missing_csv_redirects(self, client):
        r = client.post("/parse-csv", content_type="multipart/form-data")
        assert r.status_code == 303
        assert "error" in r.headers["Location"]

    def test_no_key_column_redirects(self, client):
        csv_bytes = make_csv(["Summary"], ["A story"])
        r = client.post(
            "/parse-csv",
            data={"jira_csv": (io.BytesIO(csv_bytes), "jira.csv")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 303
        assert "error" in r.headers["Location"]

    def test_keys_present(self, client):
        csv_bytes = make_csv(
            ["Issue key", "Sprint"],
            ["ALPHA-10", "Y26.SP01"],
            ["ALPHA-11", ""],
        )
        r = client.post(
            "/parse-csv",
            data={"jira_csv": (io.BytesIO(csv_bytes), "jira.csv")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        data = json.loads(r.data)
        keys = [item["key"] for item in data]
        assert "ALPHA-10" in keys
        assert "ALPHA-11" in keys


# ── Tab 3: Office – /sync with pasted JSON text ───────────────────────────────

class TestOfficeSyncPastedJson:
    def _stories_json(self):
        return json.dumps([
            {"key": "PROJ-1", "sprint": "Sprint 4"},
            {"key": "PROJ-2", "sprint": "Sprint 5"},
        ])

    def test_paste_adds_stories_to_excel(self, client):
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={
                "stories_json_text": self._stories_json(),
                "excel_file": (io.BytesIO(xl), "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        rows = read_excel_bytes(r.data)
        keys = [str(row[1]) for row in rows[1:] if row[1]]
        assert "PROJ-1" in keys
        assert "PROJ-2" in keys

    def test_paste_sprints_written_to_excel(self, client):
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={
                "stories_json_text": self._stories_json(),
                "excel_file": (io.BytesIO(xl), "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        rows = read_excel_bytes(r.data)
        sprints = [str(row[0]) for row in rows[1:] if row[0]]
        assert "Sprint 4" in sprints
        assert "Sprint 5" in sprints

    def test_invalid_json_text_redirects(self, client):
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={
                "stories_json_text": "not valid json !!!",
                "excel_file": (io.BytesIO(xl), "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 303
        assert "error" in r.headers["Location"]

    def test_missing_json_and_csv_redirects(self, client):
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={"excel_file": (io.BytesIO(xl), "tracker.xlsx")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 303
        assert "error" in r.headers["Location"]

    def test_json_file_upload_still_works(self, client):
        stories = [{"key": "PROJ-99", "sprint": "Sprint 1"}]
        json_bytes = json.dumps(stories).encode()
        xl = make_excel([["Sprint", "Issue Key"]])
        r = client.post(
            "/sync",
            data={
                "stories_json": (io.BytesIO(json_bytes), "stories.json"),
                "excel_file":   (io.BytesIO(xl), "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        rows = read_excel_bytes(r.data)
        keys = [str(row[1]) for row in rows[1:] if row[1]]
        assert "PROJ-99" in keys

    def test_existing_rows_not_duplicated(self, client):
        xl = make_excel([["Sprint", "Issue Key"], ["Sprint 4", "PROJ-1"]])
        r = client.post(
            "/sync",
            data={
                "stories_json_text": self._stories_json(),
                "excel_file": (io.BytesIO(xl), "tracker.xlsx"),
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        rows = read_excel_bytes(r.data)
        keys = [str(row[1]) for row in rows[1:] if row[1]]
        assert keys.count("PROJ-1") == 1
