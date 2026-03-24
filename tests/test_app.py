"""Tests for app.py — CSV parser, Excel sync, and Flask routes."""
import io
import json
from datetime import date

import openpyxl
import pytest

from app import app, parse_jira_csv, sync_excel


# ── Helpers ───────────────────────────────────────────────────────────────────

def csv_bytes(*lines):
    return "\n".join(lines).encode("utf-8")


def make_excel(headers, rows=None):
    """Return bytes of a minimal .xlsx with given headers and optional data rows."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in (rows or []):
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── parse_jira_csv ────────────────────────────────────────────────────────────

class TestParseJiraCsv:

    def test_basic_six_stories(self):
        data = csv_bytes(
            "Issue key,Summary,Sprint",
            "PROJ-1,Story one,Sprint 1",
            "PROJ-2,Story two,Sprint 1",
            "PROJ-3,Story three,Sprint 2",
            "PROJ-4,Story four,Sprint 2",
            "PROJ-5,Story five,Sprint 3",
            "PROJ-6,Story six,Sprint 3",
        )
        stories, warnings = parse_jira_csv(data)
        assert len(stories) == 6
        assert stories[0] == {"key": "PROJ-1", "sprint": "Sprint 1"}
        assert stories[5] == {"key": "PROJ-6", "sprint": "Sprint 3"}

    def test_filters_out_non_key_rows(self):
        """Rows like 'Product Owner', metadata, blank keys must be dropped."""
        data = csv_bytes(
            "Issue key,Summary,Sprint",
            "PROJ-1,Real story,Sprint 1",
            "Product Owner,Some metadata,",
            ",Empty key row,Sprint 1",
            "not-a-key,bad format,Sprint 1",
            "PROJ-2,Another real story,Sprint 2",
            "123,pure number,Sprint 1",
            "PROJ-3,Third story,Sprint 1",
        )
        stories, _ = parse_jira_csv(data)
        keys = [s["key"] for s in stories]
        assert keys == ["PROJ-1", "PROJ-2", "PROJ-3"]

    def test_empty_rows_skipped(self):
        data = csv_bytes(
            "Issue key,Sprint",
            "PROJ-10,Sprint 1",
            "",
            "PROJ-11,Sprint 2",
        )
        stories, _ = parse_jira_csv(data)
        assert len(stories) == 2

    def test_no_sprint_column_gives_warning(self):
        data = csv_bytes(
            "Issue key,Summary",
            "PROJ-1,Story one",
        )
        stories, warnings = parse_jira_csv(data)
        assert len(stories) == 1
        assert stories[0]["sprint"] == ""
        assert any("sprint" in w.lower() for w in warnings)

    def test_missing_key_column_raises(self):
        data = csv_bytes(
            "Summary,Status",
            "Some story,Open",
        )
        with pytest.raises(RuntimeError, match="Issue Key"):
            parse_jira_csv(data)

    def test_empty_csv_raises(self):
        with pytest.raises(RuntimeError, match="empty"):
            parse_jira_csv(b"Issue key\n")  # header only, no data

    def test_totally_empty_raises(self):
        with pytest.raises(RuntimeError):
            parse_jira_csv(b"")

    def test_alternate_header_names(self):
        """Accepts 'Key' and 'Custom field (Sprint)' variants."""
        data = csv_bytes(
            "Key,Summary,Custom field (Sprint)",
            "ABC-42,Story,Sprint 5",
        )
        stories, _ = parse_jira_csv(data)
        assert stories[0]["key"] == "ABC-42"
        assert stories[0]["sprint"] == "Sprint 5"

    def test_case_insensitive_key_matching(self):
        data = csv_bytes(
            "issue key,Sprint",
            "PROJ-99,Sprint 1",
        )
        stories, _ = parse_jira_csv(data)
        assert stories[0]["key"] == "PROJ-99"

    def test_utf8_bom_handled(self):
        raw = "Issue key,Sprint\nPROJ-1,Sprint 1\n"
        stories, _ = parse_jira_csv(raw.encode("utf-8-sig"))  # encode adds the BOM
        assert stories[0]["key"] == "PROJ-1"

    def test_key_with_underscores_and_numbers(self):
        """Project codes like 'MY_PROJ2-10' should be valid."""
        data = csv_bytes(
            "Issue key,Sprint",
            "MY_PROJ2-10,Sprint 1",
        )
        stories, _ = parse_jira_csv(data)
        assert stories[0]["key"] == "MY_PROJ2-10"

    def test_sprint_normalised(self):
        """Sprint codes like 'Y26.SP04' are normalised to 'Sprint 4'."""
        data = csv_bytes(
            "Issue key,Sprint",
            "PROJ-1,Y26.SP04",
        )
        stories, _ = parse_jira_csv(data)
        assert stories[0]["sprint"] == "Sprint 4"


# ── sync_excel ────────────────────────────────────────────────────────────────

class TestSyncExcel:

    def _read_excel_rows(self, excel_bytes):
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        return list(ws.iter_rows(values_only=True))

    def test_new_stories_appended(self):
        today_str = date.today().strftime("%-m/%-d/%Y")
        xl = make_excel(["Sprint", "Key", today_str])
        stories = [{"key": "PROJ-1", "sprint": "Sprint 1"},
                   {"key": "PROJ-2", "sprint": "Sprint 2"}]
        result, added, skipped = sync_excel(xl, stories)
        assert added == ["PROJ-1", "PROJ-2"]
        assert skipped == []
        rows = self._read_excel_rows(result)
        keys = [r[1] for r in rows[1:]]
        assert "PROJ-1" in keys
        assert "PROJ-2" in keys

    def test_existing_keys_skipped(self):
        today_str = date.today().strftime("%-m/%-d/%Y")
        xl = make_excel(["Sprint", "Key", today_str],
                        rows=[["Sprint 1", "PROJ-1", None]])
        stories = [{"key": "PROJ-1", "sprint": "Sprint 1"},
                   {"key": "PROJ-2", "sprint": "Sprint 2"}]
        _, added, skipped = sync_excel(xl, stories)
        assert added == ["PROJ-2"]
        assert skipped == ["PROJ-1"]

    def test_no_duplicates_within_batch(self):
        xl = make_excel(["Sprint", "Key"])
        stories = [{"key": "PROJ-1", "sprint": "S1"},
                   {"key": "PROJ-1", "sprint": "S1"}]  # duplicate
        _, added, skipped = sync_excel(xl, stories)
        assert added == ["PROJ-1"]
        assert skipped == ["PROJ-1"]

    def test_sprint_written_to_column_a(self):
        xl = make_excel(["Sprint", "Key"])
        stories = [{"key": "PROJ-7", "sprint": "Sprint 3"}]
        result, _, _ = sync_excel(xl, stories)
        rows = self._read_excel_rows(result)
        assert rows[1][0] == "Sprint 3"
        assert rows[1][1] == "PROJ-7"

    def test_empty_stories_list(self):
        xl = make_excel(["Sprint", "Key"])
        _, added, skipped = sync_excel(xl, [])
        assert added == []
        assert skipped == []


# ── Flask routes ──────────────────────────────────────────────────────────────

@pytest.fixture
def client():
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


class TestParseCsvRoute:

    def test_returns_json_file(self, client):
        data = csv_bytes(
            "Issue key,Sprint",
            "PROJ-1,Sprint 1",
            "PROJ-2,Sprint 2",
        )
        resp = client.post(
            "/parse-csv",
            data={"jira_csv": (io.BytesIO(data), "export.csv")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 200
        assert resp.content_type == "application/json"
        stories = json.loads(resp.data)
        assert len(stories) == 2
        assert stories[0]["key"] == "PROJ-1"

    def test_no_file_redirects_with_error(self, client):
        resp = client.post("/parse-csv", data={}, content_type="multipart/form-data")
        assert resp.status_code == 303
        assert "error" in resp.headers["Location"]

    def test_dummy_rows_excluded_from_json(self, client):
        data = csv_bytes(
            "Issue key,Sprint",
            "PROJ-1,Sprint 1",
            "Product Owner,,",
            "PROJ-2,Sprint 2",
        )
        resp = client.post(
            "/parse-csv",
            data={"jira_csv": (io.BytesIO(data), "export.csv")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 200
        stories = json.loads(resp.data)
        assert len(stories) == 2


class TestSyncRoute:

    def _make_multipart(self, stories, xl_bytes):
        return {
            "stories_json": (io.BytesIO(json.dumps(stories).encode()), "stories.json"),
            "excel_file": (io.BytesIO(xl_bytes), "tracker.xlsx"),
        }

    def test_sync_via_json(self, client):
        xl = make_excel(["Sprint", "Key"])
        stories = [{"key": "PROJ-1", "sprint": "Sprint 1"}]
        resp = client.post(
            "/sync",
            data=self._make_multipart(stories, xl),
            content_type="multipart/form-data",
        )
        assert resp.status_code == 200
        assert "spreadsheet" in resp.content_type

    def test_missing_excel_redirects(self, client):
        resp = client.post(
            "/sync",
            data={"stories_json": (io.BytesIO(b"[]"), "s.json")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 303

    def test_no_input_redirects(self, client):
        xl = make_excel(["Sprint", "Key"])
        resp = client.post(
            "/sync",
            data={"excel_file": (io.BytesIO(xl), "tracker.xlsx")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 303

    def test_fetch_returns_summary_json(self, client):
        xl = make_excel(["Sprint", "Key"],
                        rows=[["Sprint 1", "PROJ-1", None]])
        stories = [
            {"key": "PROJ-1", "sprint": "Sprint 1"},  # already exists
            {"key": "PROJ-2", "sprint": "Sprint 2"},  # new
        ]
        resp = client.post(
            "/sync",
            data=self._make_multipart(stories, xl),
            content_type="multipart/form-data",
            headers={"X-Requested-With": "fetch"},
        )
        assert resp.status_code == 200
        assert resp.content_type == "application/json"
        data = json.loads(resp.data)
        assert data["added"] == ["PROJ-2"]
        assert data["skipped"] == ["PROJ-1"]
        assert "excel_b64" in data
        assert "filename" in data


class TestDuplicateSprintColumn:

    def test_picks_column_with_higher_sprint_number(self):
        """When two Sprint columns exist, use the one with the higher sprint number."""
        data = csv_bytes(
            "Issue key,Sprint,Summary,Sprint",
            "PROJ-1,Sprint 2,Story one,Sprint 5",
            "PROJ-2,Sprint 2,Story two,Sprint 5",
        )
        stories, _ = parse_jira_csv(data)
        assert all(s["sprint"] == "Sprint 5" for s in stories)

    def test_picks_higher_even_if_earlier_column(self):
        """Column order doesn't matter — highest number wins."""
        data = csv_bytes(
            "Issue key,Sprint,Summary,Sprint",
            "PROJ-1,Sprint 9,Story one,Sprint 3",
        )
        stories, _ = parse_jira_csv(data)
        assert stories[0]["sprint"] == "Sprint 9"
