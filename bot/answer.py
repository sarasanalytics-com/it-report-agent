#!/usr/bin/env python3
"""Answer engine for the IT Slack bot.

Pulls live IT data by reusing the existing pipeline scripts (fetch-excel,
fetch-issues, generate-report), builds a grounded context from the generated
report, and asks Claude to answer the user's question using ONLY that data.

The data is cached and refreshed on a TTL so rapid-fire questions don't
re-download the spreadsheets every time.

CLI usage (for local testing once env + data are set up):
    python bot/answer.py "how many laptops are in stock?"
    python bot/answer.py --no-refresh "what did we spend on software in June?"
"""

from __future__ import annotations

import os
import sys
import time
import logging
import json
import pathlib
import threading
import subprocess

import anthropic

log = logging.getLogger("it-bot.answer")

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

# Most-complete report (monthly includes everything weekly does and more).
REPORT_TYPE = os.environ.get("BOT_REPORT_TYPE", "monthly")
# How long (seconds) to reuse fetched data before refreshing.
REFRESH_TTL = int(os.environ.get("DATA_REFRESH_TTL", "600"))
# Haiku is cheap/fast and plenty for grounded Q&A; override for higher quality.
MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-haiku-4-5")
MAX_ANSWER_TOKENS = int(os.environ.get("BOT_MAX_TOKENS", "900"))

_client: anthropic.Anthropic | None = None
_lock = threading.Lock()
_last_refresh = 0.0
_context_cache = ""
_refreshing = False

SYSTEM_PROMPT = """You are a friendly IT helper in Slack for the Head of HR, \
who is NOT technical. You answer her questions about company laptops and IT in \
simple, everyday language.

Who you're talking to:
- The Head of HR. She cares most about: are new joiners going to have a laptop \
ready on day one, how many spare laptops we have, the status of any IT requests \
raised by employees, and roughly what we're spending. Frame answers around \
people and readiness, not technical detail.

How to answer:
- Use plain, warm English. NO technical jargon or abbreviations. Avoid terms \
like "asset tag", "OS", "MDM", "procurement", "SKU", "aging buckets". If you \
must mention something technical, explain it in one short phrase.
- Be brief: open with a one-line direct answer, then at most 2-4 short "• " \
bullets only if they genuinely help. Don't pad.
- Put the key figure or name in *bold*. A tasteful emoji (✅ ⚠️ 💻 👤) is fine.
- For a new-joiner question, give a clear *Yes / Not yet* on whether a laptop \
is ready, plus the joining date.
- For a person question ("what laptop does X have", "is X's laptop old?"), use \
the EMPLOYEE LAPTOP DIRECTORY — give their laptop, its age, and whether it's \
due for replacement. If the name isn't an exact match, say who you think they \
mean or ask her to confirm the spelling.
- For accessories ("does X have a monitor/headset?"), use PERIPHERALS BY PERSON.
- For onboarding ("is X ready?", "what's pending for the new joiner?"), use \
UPCOMING JOINERS — give the joining date, the laptop needed, and which checklist \
items are still pending; judge laptop readiness by comparing the laptop needed \
against spare stock.
- For offboarding ("did X return their laptop?", "what's pending return?"), use \
LAPTOP RETURNS / OFFBOARDING.
- For her team's requests ("any IT tickets for X?", "how long pending?"), use \
IT TICKETS — match the person in 'Raised by', and quote the days open.
- For software/subscriptions ("what apps do we pay for?", "what renews next \
month?", "are we within budget?"), use SOFTWARE & LICENSES and IT BUDGET vs \
ACTUAL — list renewals due and the relevant figures.
- For money, keep the currency symbols exactly as in the data (vendor bills are \
in ₹, laptop costs in $) and don't convert them yourself.

Strict rules:
- Only use the DATA section below. If something isn't in the data, say warmly \
that you don't have that detail and suggest she check with the IT team — never \
guess or make up numbers, names, or dates.
- If she asks something unrelated to laptops/IT/people-IT-needs, gently say it's \
outside what you can see.
"""


def _client_singleton() -> anthropic.Anthropic:
    global _client
    if _client is None:
        # Reads ANTHROPIC_API_KEY from the environment.
        _client = anthropic.Anthropic()
    return _client


def _run_step(cmd: list[str], timeout: int = 240) -> bool:
    """Run a pipeline script; return True on success. Never raises."""
    try:
        subprocess.run(cmd, cwd=ROOT, env=os.environ.copy(),
                       check=True, capture_output=True, text=True, timeout=timeout)
        return True
    except Exception as exc:  # noqa: BLE001 - we want to keep going regardless
        stderr = getattr(exc, "stderr", "") or str(exc)
        log.warning("pipeline step failed: %s — %s", " ".join(cmd), str(stderr)[-400:])
        return False


def _build_context() -> str:
    parts = []
    full = OUTPUT / "full-report.md"
    directory = OUTPUT / "bot-context.md"
    metrics = OUTPUT / "metrics.json"
    if full.exists():
        parts.append("# LATEST IT REPORT\n" + full.read_text(encoding="utf-8"))
    if directory.exists():
        parts.append(directory.read_text(encoding="utf-8"))
    if metrics.exists():
        parts.append("# STRUCTURED METRICS (JSON)\n" + metrics.read_text(encoding="utf-8"))
    return "\n\n".join(parts)


def _do_refresh() -> None:
    """Re-fetch data and rebuild the cached context. Serialized by _lock."""
    global _last_refresh, _context_cache
    with _lock:
        log.info("Refreshing IT data …")
        _run_step(["python", "scripts/fetch-excel.py"])      # spreadsheets (critical)
        _run_step(["python", "scripts/fetch-issues.py"])     # ClickUp tickets (best-effort)
        _run_step(["python", "scripts/generate-report.py", REPORT_TYPE])
        ctx = _build_context()
        if ctx:
            _context_cache = ctx
            _last_refresh = time.time()
            log.info("Data refreshed (%d chars of context).", len(ctx))
        elif not _context_cache:
            _context_cache = "(no IT data is currently available)"


def warm() -> None:
    """Blocking initial load — call once at startup."""
    _do_refresh()


def get_context() -> str:
    """Return the cached context immediately. If it's missing, load it
    synchronously (first call only); if it's stale, kick off a background
    refresh but still answer from the current cache — so questions stay fast."""
    global _refreshing
    if not _context_cache:
        _do_refresh()
    elif (time.time() - _last_refresh) >= REFRESH_TTL and not _refreshing:
        _refreshing = True

        def _bg():
            global _refreshing
            try:
                _do_refresh()
            finally:
                _refreshing = False

        threading.Thread(target=_bg, daemon=True).start()
    return _context_cache or "(no IT data is currently available)"


# Back-compat alias (app startup warms via this name too).
def refresh(force: bool = False) -> str:
    _do_refresh() if force else get_context()
    return _context_cache


def get_report_blocks() -> tuple[list | None, str]:
    """Ensure data is current, then return the Block Kit report + a text
    fallback for an on-demand 'send me the report' request."""
    get_context()  # generates slack-blocks.json / slack-summary.md if needed
    blocks = None
    bpath = OUTPUT / "slack-blocks.json"
    if bpath.exists():
        try:
            blocks = json.loads(bpath.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            blocks = None
    spath = OUTPUT / "slack-summary.md"
    summary = spath.read_text(encoding="utf-8")[:2900] if spath.exists() else "IT report"
    return blocks, summary


def answer_question(question: str) -> str:
    """Answer a natural-language question from the live IT data."""
    context = get_context()
    client = _client_singleton()
    msg = client.messages.create(
        model=MODEL,
        max_tokens=MAX_ANSWER_TOKENS,
        system=SYSTEM_PROMPT + "\n\n===== DATA =====\n" + context[:150000],
        messages=[{"role": "user", "content": question}],
    )
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text").strip()
    return text or "I couldn't find an answer to that in the current IT data."


def _main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    args = sys.argv[1:]
    force = True
    if "--no-refresh" in args:
        args.remove("--no-refresh")
        # Use whatever context already exists without re-fetching.
        global _context_cache, _last_refresh
        _context_cache = _build_context() or "(no IT data is currently available)"
        _last_refresh = time.time()
        force = False
    if not args:
        print('Usage: python bot/answer.py [--no-refresh] "your question"', file=sys.stderr)
        sys.exit(1)
    question = " ".join(args)
    if force:
        refresh(force=True)
    print(answer_question(question))


if __name__ == "__main__":
    _main()
