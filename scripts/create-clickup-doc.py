#!/usr/bin/env python3
"""
Create a ClickUp task in the IT team space with the full report as its description.

We use a task (not a Doc) because the ClickUp Docs API v3 is not publicly
available for all plans. The task serves as the report container.

Usage:
    python create-clickup-doc.py <markdown_file> <doc_title>

Required environment variables:
    CLICKUP_API_TOKEN    – ClickUp personal or API token
    CLICKUP_WORKSPACE_ID – Team/workspace ID (unused, kept for compat)
    CLICKUP_SPACE_ID     – Space ID; a list within this space is auto-detected
"""

import os
import sys
import pathlib

import requests

API_TOKEN = os.environ["CLICKUP_API_TOKEN"]
SPACE_ID = os.environ.get("CLICKUP_SPACE_ID", "")
CLICKUP_API_V2 = "https://api.clickup.com/api/v2"

HEADERS = {
    "Authorization": API_TOKEN,
    "Content-Type": "application/json",
}


def get_or_create_list(list_name: str = "IT Reports") -> str:
    """Find or create a list named `list_name` in the space."""
    # Get folderless lists in the space
    url = f"{CLICKUP_API_V2}/space/{SPACE_ID}/list"
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    for lst in resp.json().get("lists", []):
        if lst["name"].lower() == list_name.lower():
            return lst["id"]

    # Create the list if it doesn't exist
    resp = requests.post(url, headers=HEADERS, json={"name": list_name}, timeout=30)
    resp.raise_for_status()
    list_id = resp.json()["id"]
    print(f"Created ClickUp list: {list_name} (ID: {list_id})")
    return list_id


def create_report_task(list_id: str, title: str, content: str) -> dict:
    """Create a task with the report content as its markdown description."""
    url = f"{CLICKUP_API_V2}/list/{list_id}/task"
    payload = {
        "name": title,
        "markdown_description": content,
        "status": "complete",
    }
    resp = requests.post(url, headers=HEADERS, json=payload, timeout=60)
    resp.raise_for_status()
    task = resp.json()
    print(f"Created ClickUp task: {title} (ID: {task['id']})")
    print(f"URL: {task.get('url', 'N/A')}")
    return task


def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: create-clickup-doc.py <markdown_file> <doc_title>", file=sys.stderr)
        sys.exit(1)

    md_path = pathlib.Path(sys.argv[1])
    title = sys.argv[2]

    if not md_path.exists():
        print(f"File not found: {md_path}", file=sys.stderr)
        sys.exit(1)

    content = md_path.read_text(encoding="utf-8").strip()
    if not content:
        print("Report file is empty — skipping ClickUp creation.", file=sys.stderr)
        sys.exit(1)

    if not SPACE_ID:
        print("CLICKUP_SPACE_ID not set — skipping ClickUp creation.", file=sys.stderr)
        sys.exit(1)

    list_id = get_or_create_list()
    create_report_task(list_id, title, content)
    print("Done.")


if __name__ == "__main__":
    main()
