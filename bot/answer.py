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
- Lead with the direct answer in a sentence, then a few short bullets if helpful.
- Keep it brief and reassuring. Use *bold* for the key number and "• " bullets.
- When she asks about a new joiner, tell her clearly whether a laptop is ready / \
being arranged, and the joining date.
- For money, keep the currency symbols exactly as in the data (vendor bills are \
in ₹, laptop costs in $) and don't convert them yourself.

Strict rule:
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
    metrics = OUTPUT / "metrics.json"
    if full.exists():
        parts.append("# LATEST IT REPORT\n" + full.read_text(encoding="utf-8"))
    if metrics.exists():
        parts.append("# STRUCTURED METRICS (JSON)\n" + metrics.read_text(encoding="utf-8"))
    return "\n\n".join(parts)


def refresh(force: bool = False) -> str:
    """Ensure data is fresh (within TTL) and return the context string."""
    global _last_refresh, _context_cache
    with _lock:
        fresh = _context_cache and (time.time() - _last_refresh) < REFRESH_TTL
        if fresh and not force:
            return _context_cache

        log.info("Refreshing IT data …")
        _run_step(["python", "scripts/fetch-excel.py"])      # spreadsheets (critical)
        _run_step(["python", "scripts/fetch-issues.py"])     # ClickUp tickets (best-effort)
        _run_step(["python", "scripts/generate-report.py", REPORT_TYPE])

        ctx = _build_context()
        if ctx:
            _context_cache = ctx
            _last_refresh = time.time()
            log.info("Data refreshed (%d chars of context).", len(ctx))
        elif _context_cache:
            log.warning("Refresh produced no new context; serving previous data.")
        else:
            _context_cache = "(no IT data is currently available)"
        return _context_cache


def answer_question(question: str) -> str:
    """Answer a natural-language question from the live IT data."""
    context = refresh()
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
