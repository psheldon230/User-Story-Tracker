import io
import csv
import json
import logging
from datetime import date, datetime

import openpyxl
from flask import Flask, jsonify, render_template, request, send_file, redirect, url_for

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ── Jira CSV parser ───────────────────────────────────────────────────────────

KEY_CANDIDATES    = ["Issue key", "Issue Key", "Key", "key"]
SPRINT_CANDIDATES = ["Sprint", "sprint", "Sprint Name", "Custom field (Sprint)"]


def _find_column(header_row: list, candidates: list):
    lower = [h.strip().lower() for h in header_row]
    for candidate in candidates:
        try:
            return lower.index(candidate.lower())
        except ValueError:
            continue
    return None


def parse_jira_csv(csv_bytes: bytes) -> tuple:
    text = csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)

    if len(rows) < 2:
        raise RuntimeError("Jira CSV appears to be empty or has no data rows.")

    header = rows[0]
    key_idx    = _find_column(header, KEY_CANDIDATES)
    sprint_idx = _find_column(header, SPRINT_CANDIDATES)

    warnings = []
    if key_idx is None:
        raise RuntimeError(
            f"Could not find an Issue Key column in the CSV. "
            f"Columns found: {', '.join(header[:10])}"
        )
    if sprint_idx is None:
        warnings.append("No Sprint column found — Sprint will be left blank.")

    stories = []
    for row in rows[1:]:
        if not row:
            continue
        key = row[key_idx].strip() if key_idx < len(row) else ""
        if not key:
            continue
        sprint = ""
        if sprint_idx is not None and sprint_idx < len(row):
            sprint = row[sprint_idx].strip()
        stories.append({"key": key, "sprint": sprint})

    return stories, warnings


# ── Excel helpers ─────────────────────────────────────────────────────────────

def _is_today_header(value) -> bool:
    today = date.today()
    if isinstance(value, datetime):
        return value.date() == today
    if isinstance(value, date):
        return value == today
    if isinstance(value, str):
        for fmt in (
            "%m/%d/%Y", "%-m/%-d/%Y", "%Y-%m-%d",
            "%m-%d-%Y", "%d/%m/%Y", "%B %d, %Y", "%b %d, %Y",
        ):
            try:
                if datetime.strptime(value.strip(), fmt).date() == today:
                    return True
            except ValueError:
                continue
    return False


def _is_test_user_header(value) -> bool:
    if not value:
        return False
    return str(value).strip().lower() in (
        "test user", "testuser", "tester", "test_user"
    )


def sync_excel(excel_bytes: bytes, stories: list) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    existing_keys: set = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > 1 and row[1] is not None:
            existing_keys.add(str(row[1]).strip())

    today_col = None
    test_user_col = None
    for cell in ws[1]:
        if cell.value is None:
            continue
        if today_col is None and _is_today_header(cell.value):
            today_col = cell.column
        if test_user_col is None and _is_test_user_header(cell.value):
            test_user_col = cell.column

    max_col = ws.max_column
    today = date.today()
    new_count = 0
    skipped_count = 0

    for story in stories:
        if story["key"] in existing_keys:
            skipped_count += 1
            continue

        new_row = [None] * max_col
        new_row[0] = story["sprint"]
        new_row[1] = story["key"]
        if today_col is not None:
            new_row[today_col - 1] = today
        if test_user_col is not None:
            new_row[test_user_col - 1] = "peter"

        ws.append(new_row)
        existing_keys.add(story["key"])
        new_count += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read(), new_count, skipped_count


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    error = request.args.get("error", "")
    return render_template("index.html", error=error)


@app.route("/parse-csv", methods=["POST"])
def parse_csv():
    """VDI step: parse Jira CSV and return a small JSON file of stories."""
    try:
        csv_file = request.files.get("jira_csv")
        if not csv_file or csv_file.filename == "":
            return redirect(url_for("index", error="Please upload your Jira CSV export."), 303)

        stories, warnings = parse_jira_csv(csv_file.read())
        if not stories:
            return redirect(url_for("index", error="No stories found in the Jira CSV."), 303)

        today_str = date.today().strftime("%Y%m%d")
        json_bytes = json.dumps(stories, indent=2).encode("utf-8")
        return send_file(
            io.BytesIO(json_bytes),
            mimetype="application/json",
            as_attachment=True,
            download_name=f"stories_{today_str}.json",
        )

    except RuntimeError as e:
        logger.warning("Handled error: %s", e)
        return redirect(url_for("index", error=str(e)), 303)
    except Exception as e:
        logger.exception("Unexpected error")
        return redirect(url_for("index", error=f"Unexpected error: {e}"), 303)


@app.route("/sync", methods=["POST"])
def sync():
    """Sync stories into Excel. Accepts either a Jira CSV or a pre-parsed JSON file."""
    try:
        csv_file   = request.files.get("jira_csv")
        json_file  = request.files.get("stories_json")
        excel_file = request.files.get("excel_file")

        if not excel_file or excel_file.filename == "":
            return redirect(url_for("index", error="Please upload your Excel tracker file."), 303)

        warnings = []
        if json_file and json_file.filename != "":
            try:
                stories = json.loads(json_file.read().decode("utf-8"))
            except Exception:
                return redirect(url_for("index", error="Could not parse the stories JSON file."), 303)
            if not stories:
                return redirect(url_for("index", error="No stories found in the JSON file."), 303)
        elif csv_file and csv_file.filename != "":
            stories, warnings = parse_jira_csv(csv_file.read())
            if not stories:
                return redirect(url_for("index", error="No stories found in the Jira CSV."), 303)
        else:
            return redirect(url_for("index", error="Please upload a Jira CSV or a stories JSON file."), 303)

        updated_bytes, new_count, skipped_count = sync_excel(excel_file.read(), stories)

        today_str = date.today().strftime("%Y%m%d")
        return send_file(
            io.BytesIO(updated_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"stories_updated_{today_str}.xlsx",
        )

    except RuntimeError as e:
        logger.warning("Handled error: %s", e)
        return redirect(url_for("index", error=str(e)), 303)
    except Exception as e:
        logger.exception("Unexpected error")
        return redirect(url_for("index", error=f"Unexpected error: {e}"), 303)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
