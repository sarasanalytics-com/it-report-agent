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

from answer import answer_question, warm, get_report_blocks, force_refresh

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("it-bot")

app = App(token=os.environ["SLACK_BOT_TOKEN"])

_MENTION_RE = re.compile(r"<@[A-Z0-9]+>")
# "send me the report", "full/weekly/it report", "report now/please", "generate report"
_REPORT_RE = re.compile(
    r"\b(full|weekly|monthly|it|latest|today'?s|the)\s+report\b"
    r"|\breport\s+(now|please)\b|send (me )?the report|generate (the )?report",
    re.I,
)
# "refresh", "reload the data", "re-fetch", "re-sync", "use the latest sheet" —
# force a fresh download after a source sheet is edited.
_REFRESH_RE = re.compile(
    r"\b(refresh|reload|re-?fetch|re-?sync|re-?download)\b"
    r"|\b(latest|updated|new|current)\s+(data|sheet|numbers|figures)\b",
    re.I,
)

WELCOME = (
    "Hi! :wave: I'm the IT helper. You can ask me about company laptops and IT "
    "in plain language — for example:\n"
    "• *Will the new joiners have laptops ready?*\n"
    "• *How many spare laptops do we have right now?*\n"
    "• *Is there a laptop ready for <name>?*\n"
    "• *What IT requests are open at the moment?*\n"
    "• *How much did we spend on software this month?*\n"
    "Just send your question and I'll take a look. :mag:\n"
    "_Tip: just edited a sheet? Say *refresh* and I'll pull the latest numbers "
    "before you ask._"
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


def _reply(text: str | None, say, client, event, thread: bool) -> None:
    question = _clean(text)
    # Thread replies under a channel @mention; reply in the main flow for DMs.
    thread_ts = (event.get("thread_ts") or event.get("ts")) if thread else None
    log.info("Question from %s (%s): %r",
             event.get("user"), event.get("channel_type") or "channel", (question or "")[:120])
    if not question or question.lower() in ("hi", "hello", "hey", "help", "?"):
        say(text=WELCOME, thread_ts=thread_ts)
        return

    # Manual data reload: pull the sheets again right now (e.g. after an edit).
    # Synchronous so the confirmation only lands once fresh data is in place.
    if _REFRESH_RE.search(question):
        say(text=":arrows_counterclockwise: Reloading the latest IT data — one moment…",
            thread_ts=thread_ts)
        try:
            force_refresh()
            say(text=":white_check_mark: Done — I've pulled the latest sheets. Ask me "
                     "your question again and I'll use the updated numbers.",
                thread_ts=thread_ts)
        except Exception as exc:  # noqa: BLE001
            log.exception("Manual refresh failed")
            say(text=f":warning: I couldn't reload the data just now ({exc}). "
                     "Please try again in a moment.", thread_ts=thread_ts)
        return

    # On-demand full report: "send me the IT report", "weekly report", etc.
    if _REPORT_RE.search(question):
        say(text=":bar_chart: Pulling the latest IT report…", thread_ts=thread_ts)
        try:
            blocks, summary = get_report_blocks()
            say(text=summary or "Here's the latest IT report:", blocks=blocks, thread_ts=thread_ts)
        except Exception as exc:  # noqa: BLE001
            log.exception("Failed to build on-demand report")
            say(text=f":warning: Couldn't build the report just now ({exc}).", thread_ts=thread_ts)
        return

    # Post an instant placeholder so it never looks dead, then edit in the answer.
    placeholder = say(text=":mag: Looking that up…", thread_ts=thread_ts)
    try:
        answer = answer_question(question)
    except Exception as exc:  # noqa: BLE001
        log.exception("Failed to answer question")
        answer = (":warning: Sorry, something went wrong on my side and I "
                  "couldn't get that answer. Please try again in a moment, or "
                  f"check with the IT team. _(details: {exc})_")
    try:
        client.chat_update(channel=placeholder["channel"], ts=placeholder["ts"], text=answer)
    except Exception:  # noqa: BLE001 - fall back to a fresh message
        say(text=answer, thread_ts=thread_ts)


def _guard(event, say) -> bool:
    """Enforce the allowlist; reply with a polite decline if not authorized."""
    if _authorized(event):
        return True
    log.info("Denied request from user %s (not on allowlist)", event.get("user"))
    say(text=DENY_MESSAGE, thread_ts=event.get("thread_ts") or event.get("ts"))
    return False


@app.event("app_mention")
def handle_mention(event, say, client):
    """Someone @mentioned the bot in a channel."""
    log.info("app_mention from %s in %s", event.get("user"), event.get("channel"))
    if not _guard(event, say):
        return
    _reply(event.get("text"), say, client, event, thread=True)


@app.event("message")
def handle_message(event, say, client):
    """Direct message to the bot. Ignore channel messages, edits and bots."""
    if event.get("channel_type") != "im":
        return
    if event.get("bot_id") or event.get("subtype"):
        return
    log.info("DM from %s", event.get("user"))
    if not _guard(event, say):
        return
    _reply(event.get("text"), say, client, event, thread=False)


def main() -> None:
    # Warm the data cache on startup so the first question is fast.
    try:
        warm()
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
