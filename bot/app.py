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


def _load_allowlist() -> set[str]:
    """Slack user IDs allowed to use the bot (from BOT_ALLOWED_USERS, comma or
    semicolon separated). Empty set means the bot is open to everyone."""
    raw = os.environ.get("BOT_ALLOWED_USERS", "")
    return {u.strip().upper() for u in raw.replace(";", ",").split(",") if u.strip()}


ALLOWED_USERS = _load_allowlist()

DENY_MESSAGE = (
    "Sorry, I can only share IT information with approved people. If you think "
    "you should have access, please contact the IT team. :lock:"
)


def _authorized(event) -> bool:
    """True when the allowlist is empty (open) or the sender is on it."""
    if not ALLOWED_USERS:
        return True
    return (event.get("user") or "").upper() in ALLOWED_USERS


def _reply(text: str | None, say, event, thread: bool) -> None:
    question = _clean(text)
    # Thread replies under a channel @mention; reply in the main flow for DMs.
    thread_ts = (event.get("thread_ts") or event.get("ts")) if thread else None
    log.info("Question from %s (%s): %r",
             event.get("user"), event.get("channel_type") or "channel", (question or "")[:120])
    if not question or question.lower() in ("hi", "hello", "hey", "help", "?"):
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


def _guard(event, say) -> bool:
    """Enforce the allowlist; reply with a polite decline if not authorized."""
    if _authorized(event):
        return True
    log.info("Denied request from user %s (not on allowlist)", event.get("user"))
    say(text=DENY_MESSAGE, thread_ts=event.get("thread_ts") or event.get("ts"))
    return False


@app.event("app_mention")
def handle_mention(event, say):
    """Someone @mentioned the bot in a channel."""
    log.info("app_mention from %s in %s", event.get("user"), event.get("channel"))
    if not _guard(event, say):
        return
    _reply(event.get("text"), say, event, thread=True)


@app.event("message")
def handle_message(event, say):
    """Direct message to the bot. Ignore channel messages, edits and bots."""
    if event.get("channel_type") != "im":
        return
    if event.get("bot_id") or event.get("subtype"):
        return
    log.info("DM from %s", event.get("user"))
    if not _guard(event, say):
        return
    _reply(event.get("text"), say, event, thread=False)


def main() -> None:
    # Warm the data cache on startup so the first question is fast.
    try:
        refresh(force=True)
    except Exception:  # noqa: BLE001
        log.warning("Initial data warm-up failed; will retry on first question.")
    if ALLOWED_USERS:
        log.info("Access restricted to %d allowlisted user(s).", len(ALLOWED_USERS))
    else:
        log.warning("BOT_ALLOWED_USERS is empty — the bot will answer ANYONE who "
                    "can DM or @mention it. Set BOT_ALLOWED_USERS to restrict access.")
    log.info("IT bot starting (Socket Mode) …")
    SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"]).start()


if __name__ == "__main__":
    main()
