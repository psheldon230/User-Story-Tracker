import io
import re
import logging
from datetime import date, datetime

import requests
import openpyxl
from flask import Flask, jsonify, render_template, request, send_file

# Suppress SSL warnings — common in corporate VDI environments with self-signed certs
try:
    from requests.packages.urllib3.exceptions import InsecureRequestWarning
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
except Exception:
    pass

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB max upload

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ── Jira helpers ──────────────────────────────────────────────────────────────

def fetch_jira_stories(jira_url: str, jql: str, auth, headers: dict) -> list:
    """Fetch all matching issues from Jira, handling pagination automatically."""
    stories = []
    start_at = 0
    max_results = 100

    while True:
        params = {
            "jql": jql,
            "startAt": start_at,
            "maxResults": max_results,
            # customfield_10020 is the standard Sprint field in Jira Server/Cloud
            "fields": "summary,status,assignee,priority,customfield_10020",
        }
        try:
            resp = requests.get(
                f"{jira_url}/rest/api/2/search",
                params=params,
                auth=auth,
                headers=headers,
                verify=False,   # corporate certs — disable SSL verification
                timeout=30,
            )
            resp.raise_for_status()
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response else "unknown"
            if status == 401:
                raise RuntimeError("Authentication failed — check your credentials or token.")
            if status == 403:
                raise RuntimeError("Access denied — your account may not have permission to run this query.")
            if status == 400:
                detail = ""
                try:
                    detail = e.response.json().get("errorMessages", [e.response.text])[0]
                except Exception:
                    detail = e.response.text
                raise RuntimeError(f"Bad Jira query: {detail}")
            raise RuntimeError(f"Jira API error {status}: {e}")
        except requests.exceptions.ConnectionError:
            raise RuntimeError(f"Could not connect to Jira at {jira_url} — check the URL and your VPN.")
        except requests.exceptions.Timeout:
            raise RuntimeError("Jira request timed out — the server may be slow or unreachable.")

        data = resp.json()
        issues = data.get("issues", [])

        for issue in issues:
            fields = issue.get("fields", {})
            sprint_name = _extract_sprint(fields.get("customfield_10020"))
            stories.append({
                "key": issue["key"],
                "sprint": sprint_name,
                "summary": fields.get("summary", ""),
            })

        start_at += len(issues)
        if start_at >= data.get("total", 0) or not issues:
            break

    return stories


def _extract_sprint(sprint_field) -> str:
    """
    Jira returns Sprint data in customfield_10020 in several formats depending
    on the version — handle all of them.
    """
    if not sprint_field:
        return ""
    if isinstance(sprint_field, list) and sprint_field:
        # Use the last entry (most recent / active sprint)
        sprint_info = sprint_field[-1]
        if isinstance(sprint_info, dict):
            return sprint_info.get("name", "")
        if isinstance(sprint_info, str):
            # Older Jira Server returns a string like:
            # "com.atlassian.greenhopper.service.sprint.Sprint@...name=Sprint 5,..."
            m = re.search(r"name=([^,\]]+)", sprint_info)
            return m.group(1).strip() if m else ""
    if isinstance(sprint_field, dict):
        return sprint_field.get("name", "")
    return ""


# ── Excel helpers ─────────────────────────────────────────────────────────────

def _is_today_header(value) -> bool:
    """Return True if a cell header value represents today's date."""
    today = date.today()
    if isinstance(value, datetime):
        return value.date() == today
    if isinstance(value, date):
        return value == today
    if isinstance(value, str):
        # Try common date string formats
        for fmt in (
            "%m/%d/%Y",   # 03/24/2026
            "%-m/%-d/%Y", # 3/24/2026  (Linux only — will fall through on Windows)
            "%#m/%#d/%Y", # 3/24/2026  (Windows strptime)
            "%Y-%m-%d",   # 2026-03-24
            "%m-%d-%Y",   # 03-24-2026
            "%d/%m/%Y",   # 24/03/2026
            "%B %d, %Y",  # March 24, 2026
            "%b %d, %Y",  # Mar 24, 2026
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
        "test user", "testuser", "tester", "test_user", "qa user", "qa_user"
    )


def sync_stories(excel_bytes: bytes, stories: list) -> tuple:
    """
    Open the workbook, find new stories not already in column B,
    and append them. Returns (updated_excel_bytes, new_count, total_jira_count).
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    # Collect existing Issue Keys from column B (index 1), starting row 2
    existing_keys: set = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > 1 and row[1] is not None:
            existing_keys.add(str(row[1]).strip())

    # Scan header row (row 1) for today's date column and test-user column
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

    for story in stories:
        if story["key"] in existing_keys:
            continue

        # Build a row with None placeholders then fill in the known positions
        new_row = [None] * max_col
        new_row[0] = story["sprint"]   # Column A — Sprint
        new_row[1] = story["key"]      # Column B — Issue Key
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
    return output.read(), new_count


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/sync", methods=["POST"])
def sync():
    try:
        # ── Inputs ──────────────────────────────────────────────
        excel_file = request.files.get("excel_file")
        if not excel_file or excel_file.filename == "":
            return jsonify({"error": "Please upload an Excel file."}), 400

        jira_url = request.form.get("jira_url", "").strip().rstrip("/")
        auth_type = request.form.get("auth_type", "pat")
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        token = request.form.get("token", "").strip()
        project_key = request.form.get("project_key", "").strip()
        filter_id = request.form.get("filter_id", "").strip()

        if not jira_url:
            return jsonify({"error": "Jira URL is required."}), 400

        # ── Auth ─────────────────────────────────────────────────
        if auth_type == "pat":
            if not token:
                return jsonify({"error": "Personal Access Token is required."}), 400
            auth = None
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            }
        else:
            if not username or not password:
                return jsonify({"error": "Username and password are required for Basic Auth."}), 400
            auth = (username, password)
            headers = {"Content-Type": "application/json"}

        # ── JQL ──────────────────────────────────────────────────
        if filter_id:
            jql = f"filter={filter_id}"
        elif project_key:
            jql = f"project={project_key} ORDER BY created DESC"
        else:
            return jsonify({"error": "Provide at least a Project Key or Filter ID."}), 400

        # ── Fetch Jira ───────────────────────────────────────────
        stories = fetch_jira_stories(jira_url, jql, auth, headers)

        # ── Sync Excel ───────────────────────────────────────────
        excel_bytes = excel_file.read()
        updated_bytes, new_count = sync_stories(excel_bytes, stories)

        today_str = date.today().strftime("%Y%m%d")
        filename = f"stories_updated_{today_str}.xlsx"

        response = send_file(
            io.BytesIO(updated_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )
        # Expose custom headers so the frontend JS can read them
        response.headers["X-New-Stories"] = str(new_count)
        response.headers["X-Total-Stories"] = str(len(stories))
        response.headers["Access-Control-Expose-Headers"] = "X-New-Stories, X-Total-Stories"
        return response

    except RuntimeError as e:
        logger.warning("Handled error: %s", e)
        return jsonify({"error": str(e)}), 502
    except Exception as e:
        logger.exception("Unexpected error during sync")
        return jsonify({"error": f"Unexpected error: {e}"}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
