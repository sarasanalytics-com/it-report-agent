#!/usr/bin/env python3
"""
Fetch IT helpdesk tickets from the ClickUp IT ticket list and write them to
data/it_issues.xlsx (sheet "IT Issues") for the report generator to consume.

The ClickUp list is the one behind the IT ticket space view, e.g.
    https://app.clickup.com/3369097/v/li/901612104806
where 901612104806 is the List ID.

Each ClickUp task becomes one issue row:
    Date Raised | Issue | Raised By | Priority | Status | Owner

Required environment variables:
    CLICKUP_API_TOKEN          – ClickUp personal/API token (same one used by
                                 create-clickup-doc.py)

Optional environment variables:
    CLICKUP_IT_ISSUES_LIST_ID  – List ID for IT tickets (default: 901612104806)

This step is best-effort: if the token is missing or the API call fails it
prints a warning and exits 0 without writing a file, so the weekly/monthly
report still generates (it then shows the "no source connected" placeholder).
"""

import datetime as dt
import os
import pathlib
import sys

import openpyxl
import requests

API_TOKEN = os.environ.get("CLICKUP_API_TOKEN", "").strip()
# Use `or` rather than a get() default so a secret that is *set but empty*
# (the case when CLICKUP_IT_ISSUES_LIST_ID isn't configured) still falls back
# to the IT ticket list rather than producing an empty list ID.
DEFAULT_LIST_ID = "901612104806"
LIST_ID = (os.environ.get("CLICKUP_IT_ISSUES_LIST_ID") or "").strip() or DEFAULT_LIST_ID
CLICKUP_API_V2 = "https://api.clickup.com/api/v2"
DATA_DIR = pathlib.Path(__file__).resolve().parent.parent / "data"

HEADERS = {"Authorization": API_TOKEN}

# ClickUp custom-field names that may hold the requester, if the workspace uses
# one instead of relying on the task creator.
REQUESTER_FIELD_NAMES = ("raised by", "reported by", "requester", "requested by", "employee")


def _ms_to_date(ms) -> str:
    """Convert a ClickUp millisecond timestamp (string/int) to YYYY-MM-DD."""
    try:
        return dt.datetime.fromtimestamp(int(ms) / 1000).strftime("%Y-%m-%d")
    except (TypeError, ValueError, OSError):
        return ""


def _requester_from_custom_fields(task: dict) -> str:
    for field in task.get("custom_fields", []) or []:
        name = str(field.get("name", "")).strip().lower()
        if name in REQUESTER_FIELD_NAMES:
            val = field.get("value")
            if isinstance(val, list) and val:
                # People-type field → list of user dicts
                first = val[0]
                if isinstance(first, dict):
                    return str(first.get("username") or first.get("email") or "").strip()
            if val:
                return str(val).strip()
    return ""


def _normalise_task(task: dict) -> dict:
    status = (task.get("status") or {}).get("status", "")
    priority_obj = task.get("priority") or {}
    priority = priority_obj.get("priority", "") if isinstance(priority_obj, dict) else ""
    creator = task.get("creator") or {}
    assignees = task.get("assignees") or []
    owner = ""
    if assignees:
        owner = str(assignees[0].get("username") or assignees[0].get("email") or "").strip()
    raised_by = _requester_from_custom_fields(task) or str(creator.get("username", "")).strip()
    return {
        "_task_id": str(task.get("id", "")),  # internal, not written to the sheet
        "Date Raised": _ms_to_date(task.get("date_created")),
        "Issue": str(task.get("name", "")).strip(),
        "Raised By": raised_by,
        "Priority": str(priority or "").strip().title(),
        "Status": str(status or "").strip().title(),
        "Owner": owner,
        "Latest Update": "",  # filled from the task's latest comment (open tickets)
    }


# Automated/template acknowledgement comments to skip when looking for the real
# "why pending" update on a ticket.
_AUTO_COMMENT_MARKERS = (
    "request has been logged",
    "thanks for reaching out",
    "your request has been",
    "auto-generated",
    "this is an automated",
)


def fetch_latest_comment(task_id: str) -> str:
    """Return the most recent meaningful comment on a task (skipping automated
    acknowledgement templates). Empty string if none / on error."""
    if not task_id:
        return ""
    url = f"{CLICKUP_API_V2}/task/{task_id}/comment"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        resp.raise_for_status()
        comments = resp.json().get("comments", [])  # ClickUp returns newest-first
    except requests.RequestException:
        return ""

    def _text(c: dict) -> str:
        # Collapse newlines so multi-line updates fit one table cell
        return " ".join((c.get("comment_text") or "").split()).strip()

    # Prefer the newest non-automated comment; fall back to newest non-empty.
    fallback = ""
    for c in comments:
        text = _text(c)
        if not text:
            continue
        if not fallback:
            fallback = text
        if not any(m in text.lower() for m in _AUTO_COMMENT_MARKERS):
            return text
    return fallback


def fetch_tasks(list_id: str) -> list[dict]:
    """Fetch all tasks (open + closed) from a ClickUp list, with pagination."""
    tasks: list[dict] = []
    page = 0
    while True:
        url = f"{CLICKUP_API_V2}/list/{list_id}/task"
        params = {
            "page": page,
            "include_closed": "true",
            "subtasks": "true",
            "order_by": "created",
            "reverse": "true",  # newest first
        }
        resp = requests.get(url, headers=HEADERS, params=params, timeout=60)
        resp.raise_for_status()
        body = resp.json()
        batch = body.get("tasks", [])
        tasks.extend(batch)
        if body.get("last_page") or not batch:
            break
        page += 1
        if page > 50:  # safety valve against runaway pagination
            break
    return tasks


def write_xlsx(rows: list[dict], dest: pathlib.Path) -> None:
    headers = ["Date Raised", "Issue", "Raised By", "Priority", "Status", "Owner", "Latest Update"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IT Issues"
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)


def main() -> None:
    dest = DATA_DIR / "it_issues.xlsx"

    if not API_TOKEN:
        print("CLICKUP_API_TOKEN not set — skipping IT issue fetch "
              "(report will show the placeholder).", file=sys.stderr)
        return  # exit 0: non-fatal

    print(f"Fetching IT tickets from ClickUp list {LIST_ID} …")
    try:
        tasks = fetch_tasks(LIST_ID)
    except requests.RequestException as exc:
        print(f"  ✗ ClickUp fetch failed: {exc}\n  Skipping (report will show "
              f"the placeholder).", file=sys.stderr)
        return  # exit 0: non-fatal — don't break the report run

    rows = [_normalise_task(t) for t in tasks]

    def _is_open(status: str) -> bool:
        return status.lower() not in (
            "resolved", "closed", "done", "complete", "completed",
            "cancelled", "canceled", "duplicate", "rejected")

    open_rows = [r for r in rows if _is_open(r["Status"])]

    # Pull the latest activity/comment for OPEN tickets only — that's the real
    # "why pending" note. Limited to open tickets to keep the API calls bounded.
    for r in open_rows:
        r["Latest Update"] = fetch_latest_comment(r.get("_task_id", ""))

    write_xlsx(rows, dest)

    print(f"  ✓ Wrote {len(rows)} ticket(s) → {dest.name} "
          f"({len(open_rows)} open, {len(rows) - len(open_rows)} resolved)")
    if open_rows:
        print(f"  Open tickets ({len(open_rows)}):")
        for r in open_rows[:50]:
            date = r["Date Raised"] or "—"
            prio = r["Priority"] or "—"
            owner = r["Owner"] or "unassigned"
            update = r.get("Latest Update") or "(no comment)"
            print(f"    • [{prio}] {r['Issue']} — {r['Status']} "
                  f"(raised {date} by {r['Raised By'] or '—'}, owner: {owner})")
            print(f"        ↳ why pending: {update}")
        if len(open_rows) > 50:
            print(f"    …and {len(open_rows) - 50} more")


if __name__ == "__main__":
    main()
