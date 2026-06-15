#!/usr/bin/env python3
"""Saras IT Slack bot (Socket Mode).

Answers the Head of HR's questions about company laptops, new-joiner readiness,
IT requests and spend — in plain language — by reading the live IT data and
asking Claude (see answer.py).

Runs as an always-on service. It connects to Slack over Socket Mode, so no
public URL or inbound webhook is required.

Required environment variables:
    SLACK_BOT_TOKEN    xoxb-… bot token (chat:write, app_mentions:read, im:*)
    SLACK_APP_TOKEN    xapp-… app-level token with connections:write (Socket Mode)
    ANTHROPIC_API_KEY  Claude API key
    + the existing pipeline secrets (Azure/SharePoint, ClickUp) so it can fetch
      live data — see bot/.env.example.
"""

from __future__ import annotations

import os
import re
import logging

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

from answer import answer_question, refresh

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("it-bot")

app = App(token=os.environ["SLACK_BOT_TOKEN"])

_MENTION_RE = re.compile(r"<@[A-Z0-9]+>")

WELCOME = (
    "Hi! :wave: I'm the IT helper. You can ask me about company laptops and IT "
    "in plain language — for example:\n"
    "• *Will the new joiners have laptops ready?*\n"
    "• *How many spare laptops do we have right now?*\n"
    "• *Is there a laptop ready for <name>?*\n"
    "• *What IT requests are open at the moment?*\n"
    "• *How much did we spend on software this month?*\n"
    "Just send your question and I'll take a look. :mag:"
)


def _clean(text: str | None) -> str:
    return _MENTION_RE.sub("", text or "").strip()


def _reply(text: str | None, say, event) -> None:
    question = _clean(text)
    thread_ts = event.get("thread_ts") or event.get("ts")
    if not question:
        say(text=WELCOME, thread_ts=thread_ts)
        return
    if question.lower() in ("hi", "hello", "hey", "help", "?"):
        say(text=WELCOME, thread_ts=thread_ts)
        return
    try:
        answer = answer_question(question)
    except Exception as exc:  # noqa: BLE001
        log.exception("Failed to answer question")
        answer = (":warning: Sorry, something went wrong on my side and I "
                  "couldn't get that answer. Please try again in a moment, or "
                  f"check with the IT team. _(details: {exc})_")
    say(text=answer, thread_ts=thread_ts)


@app.event("app_mention")
def handle_mention(event, say):
    """Someone @mentioned the bot in a channel."""
    _reply(event.get("text"), say, event)


@app.event("message")
def handle_message(event, say):
    """Direct message to the bot. Ignore channel messages, edits and bots."""
    if event.get("channel_type") != "im":
        return
    if event.get("bot_id") or event.get("subtype"):
        return
    _reply(event.get("text"), say, event)


def main() -> None:
    # Warm the data cache on startup so the first question is fast.
    try:
        refresh(force=True)
    except Exception:  # noqa: BLE001
        log.warning("Initial data warm-up failed; will retry on first question.")
    log.info("IT bot starting (Socket Mode) …")
    SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"]).start()


if __name__ == "__main__":
    main()
