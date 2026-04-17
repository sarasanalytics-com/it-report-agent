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


def _read_and_trim(path: pathlib.Path) -> str:
    text = path.read_text(encoding="utf-8").strip()
    if len(text) > 3900:
        text = text[:3900] + "\n\n_… message truncated._"
    return text


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: post-to-slack.py <summary_md> [docx_file] [--alert <alert_md>]",
              file=sys.stderr)
        sys.exit(1)

    args = sys.argv[1:]
    alert_path = None
    if "--alert" in args:
        i = args.index("--alert")
        if i + 1 >= len(args):
            print("--alert flag requires a file path", file=sys.stderr)
            sys.exit(1)
        alert_path = pathlib.Path(args[i + 1])
        args = args[:i] + args[i + 2:]

    md_path = pathlib.Path(args[0])
    docx_path = pathlib.Path(args[1]) if len(args) >= 2 else None

    if not md_path.exists():
        print(f"File not found: {md_path}", file=sys.stderr)
        sys.exit(1)

    content = _read_and_trim(md_path)
    if not content:
        print("Report file is empty — skipping Slack post.", file=sys.stderr)
        sys.exit(1)

    # Post the summary message
    ts = slack_post(content)

    # Upload the Word doc in a thread under the summary
    if docx_path and docx_path.exists():
        title = docx_path.stem.replace("-", " ").replace("_", " ")
        slack_upload_file(docx_path, title, ts)
    elif docx_path:
        print(f"Warning: docx file not found: {docx_path}", file=sys.stderr)

    # Post the alert as a separate standalone message
    if alert_path and alert_path.exists():
        alert_text = _read_and_trim(alert_path)
        if alert_text:
            slack_post(alert_text)
    elif alert_path:
        print(f"Warning: alert file not found: {alert_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
