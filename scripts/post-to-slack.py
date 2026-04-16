#!/usr/bin/env python3
"""
Post a Markdown report summary to a Slack channel.

Usage:
    python post-to-slack.py <markdown_file>

Required environment variables:
    SLACK_BOT_TOKEN   – xoxb- bot token with chat:write scope
    SLACK_CHANNEL     – channel ID or name (e.g., C0123456789 or #it-reports)
"""

import os
import sys
import pathlib

import requests


SLACK_TOKEN = os.environ["SLACK_BOT_TOKEN"]
CHANNEL = os.environ.get("SLACK_CHANNEL", "#it-reports")
SLACK_POST_URL = "https://slack.com/api/chat.postMessage"


def slack_post(text: str) -> None:
    """Post a message to Slack using chat.postMessage."""
    resp = requests.post(
        SLACK_POST_URL,
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
    print(f"Posted to {CHANNEL} (ts: {body['ts']})")


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: post-to-slack.py <markdown_file>", file=sys.stderr)
        sys.exit(1)

    md_path = pathlib.Path(sys.argv[1])
    if not md_path.exists():
        print(f"File not found: {md_path}", file=sys.stderr)
        sys.exit(1)

    content = md_path.read_text(encoding="utf-8").strip()
    if not content:
        print("Report file is empty — skipping Slack post.", file=sys.stderr)
        sys.exit(1)

    # Slack has a 4000-char limit per message; truncate if needed
    if len(content) > 3900:
        content = content[:3900] + "\n\n_… report truncated. See full report in ClickUp._"

    slack_post(content)


if __name__ == "__main__":
    main()
