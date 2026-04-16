#!/usr/bin/env python3
"""
Create a ClickUp Doc in the IT team workspace with the full report content.

Usage:
    python create-clickup-doc.py <markdown_file> <doc_title>

Required environment variables:
    CLICKUP_API_TOKEN    – ClickUp personal or API token
    CLICKUP_WORKSPACE_ID – Team/workspace ID
    CLICKUP_SPACE_ID     – Space ID where docs should be created
"""

import os
import sys
import pathlib

import requests

API_TOKEN = os.environ["CLICKUP_API_TOKEN"]
WORKSPACE_ID = os.environ["CLICKUP_WORKSPACE_ID"]
SPACE_ID = os.environ.get("CLICKUP_SPACE_ID", "")
CLICKUP_API_BASE = "https://api.clickup.com/api/v3"


def create_doc(title: str, content: str) -> dict:
    """Create a ClickUp Doc via API v3."""
    url = f"{CLICKUP_API_BASE}/workspaces/{WORKSPACE_ID}/docs"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    payload = {
        "name": title,
        "parent": {"id": SPACE_ID, "type": 5} if SPACE_ID else None,
    }
    # Remove None values
    payload = {k: v for k, v in payload.items() if v is not None}

    resp = requests.post(url, headers=headers, json=payload, timeout=30)
    resp.raise_for_status()
    doc = resp.json()
    doc_id = doc["id"]
    print(f"Created ClickUp doc: {title} (ID: {doc_id})")

    # Add content as a page
    page_url = f"{CLICKUP_API_BASE}/workspaces/{WORKSPACE_ID}/docs/{doc_id}/pages"
    page_payload = {
        "name": title,
        "content": content,
        "content_format": "text/md",
    }
    page_resp = requests.post(page_url, headers=headers, json=page_payload, timeout=30)
    page_resp.raise_for_status()
    print(f"Added report content to doc page.")

    return doc


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
        print("Report file is empty — skipping ClickUp doc creation.", file=sys.stderr)
        sys.exit(1)

    create_doc(title, content)
    print("Done.")


if __name__ == "__main__":
    main()
