#!/usr/bin/env python3
"""
Post a Markdown report summary to Slack and optionally upload a Word doc.

Usage:
    python post-to-slack.py <summary_md> [docx_file]

Required environment variables:
    SLACK_BOT_TOKEN   – xoxb- bot token with chat:write and files:write scopes
    SLACK_CHANNEL     – channel ID or name (e.g., C0123456789 or #it-reports)
"""

import os
import sys
import pathlib

import requests


SLACK_TOKEN = os.environ["SLACK_BOT_TOKEN"]
CHANNEL = os.environ.get("SLACK_CHANNEL", "#it-reports")


def slack_post(text: str) -> str:
    """Post a message to Slack. Returns the message timestamp."""
    resp = requests.post(
        "https://slack.com/api/chat.postMessage",
        headers={"Authorization": f"Bearer {SLACK_TOKEN}"},
        json={
            "channel": CHANNEL,
            "text": text,
            "mrkdwn": True,
        },
        timeout=30,
    )
    resp.raise_for_status()
    body = resp.json()
    if not body.get("ok"):
        print(f"Slack API error: {body.get('error', body)}", file=sys.stderr)
        sys.exit(1)
    ts = body["ts"]
    print(f"Posted summary to {CHANNEL} (ts: {ts})")
    return ts


def slack_upload_file(file_path: pathlib.Path, title: str, thread_ts: str) -> None:
    """Upload a file to Slack in a thread under the summary message."""
    # Step 1: Get upload URL
    resp = requests.post(
        "https://slack.com/api/files.getUploadURLExternal",
        headers={"Authorization": f"Bearer {SLACK_TOKEN}"},
        data={
            "filename": file_path.name,
            "length": file_path.stat().st_size,
        },
        timeout=30,
    )
    resp.raise_for_status()
    body = resp.json()
    if not body.get("ok"):
        print(f"Slack upload URL error: {body.get('error', body)}", file=sys.stderr)
        sys.exit(1)

    upload_url = body["upload_url"]
    file_id = body["file_id"]

    # Step 2: Upload the file content
    with open(file_path, "rb") as f:
        resp = requests.post(upload_url, files={"file": f}, timeout=60)
        resp.raise_for_status()

    # Step 3: Complete the upload and share to channel
    resp = requests.post(
        "https://slack.com/api/files.completeUploadExternal",
        headers={
            "Authorization": f"Bearer {SLACK_TOKEN}",
            "Content-Type": "application/json",
        },
        json={
            "files": [{"id": file_id, "title": title}],
            "channel_id": CHANNEL,
            "thread_ts": thread_ts,
            "initial_comment": f"📎 Full report attached: *{title}*",
        },
        timeout=30,
    )
    resp.raise_for_status()
    body = resp.json()
    if not body.get("ok"):
        print(f"Slack complete upload error: {body.get('error', body)}", file=sys.stderr)
        sys.exit(1)
    print(f"Uploaded {file_path.name} to {CHANNEL}")


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: post-to-slack.py <summary_md> [docx_file]", file=sys.stderr)
        sys.exit(1)

    md_path = pathlib.Path(sys.argv[1])
    docx_path = pathlib.Path(sys.argv[2]) if len(sys.argv) >= 3 else None

    if not md_path.exists():
        print(f"File not found: {md_path}", file=sys.stderr)
        sys.exit(1)

    content = md_path.read_text(encoding="utf-8").strip()
    if not content:
        print("Report file is empty — skipping Slack post.", file=sys.stderr)
        sys.exit(1)

    # Slack has a 4000-char limit per message; truncate if needed
    if len(content) > 3900:
        content = content[:3900] + "\n\n_… summary truncated. Full report attached below._"

    # Post the summary message
    ts = slack_post(content)

    # Upload the Word doc in a thread if provided
    if docx_path and docx_path.exists():
        title = docx_path.stem.replace("-", " ").replace("_", " ")
        slack_upload_file(docx_path, title, ts)
    elif docx_path:
        print(f"Warning: docx file not found: {docx_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
