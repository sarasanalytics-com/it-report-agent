#!/usr/bin/env python3
"""Question log + a tiny built-in web page to review it.

Every question asked to the IT bot is appended to a JSONL file; a lightweight
HTTP server (stdlib only, no extra deps) serves it as a readable web page at
``/`` and as a CSV download at ``/questions.csv``.

Persistence: the log lives at ``QUESTION_LOG_DIR/question-log.jsonl``
(default: ``<repo>/output``). On Render the container filesystem is ephemeral,
so to keep history across deploys set ``QUESTION_LOG_DIR`` to a mounted
persistent disk (e.g. ``/var/data``).

Access: if ``LOG_VIEW_TOKEN`` is set, the page requires ``?token=<value>``;
otherwise it's open (a warning is logged and shown on the page).

The server binds ``PORT`` (Render sets this for a Web Service; default 3000).
"""

from __future__ import annotations

import os
import io
import csv
import json
import html
import logging
import pathlib
import threading
import datetime as dt
from urllib.parse import urlparse, parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

log = logging.getLogger("it-bot.qlog")

ROOT = pathlib.Path(__file__).resolve().parent.parent
LOG_DIR = pathlib.Path(os.environ.get("QUESTION_LOG_DIR", str(ROOT / "output")))
LOG_PATH = LOG_DIR / "question-log.jsonl"
VIEW_TOKEN = os.environ.get("LOG_VIEW_TOKEN", "").strip()
MAX_ROWS = int(os.environ.get("LOG_VIEW_MAX_ROWS", "5000"))

_lock = threading.Lock()


def log_question(user: str | None, question: str,
                 channel_type: str = "", user_name: str = "") -> None:
    """Append one question to the log. Never raises (logging must not break a reply)."""
    if not (question or "").strip():
        return
    rec = {
        "ts": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "user": user or "",
        "user_name": user_name or "",
        "channel_type": channel_type or "",
        "question": question.strip(),
    }
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with _lock, open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:  # noqa: BLE001
        log.exception("could not write question log")


def read_questions(limit: int = MAX_ROWS) -> list[dict]:
    """Return logged questions, newest first (capped at `limit`)."""
    if not LOG_PATH.exists():
        return []
    rows: list[dict] = []
    try:
        with _lock, open(LOG_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    except OSError:
        log.exception("could not read question log")
        return []
    rows.reverse()
    return rows[:limit]


def _fmt_ts(iso: str) -> str:
    try:
        d = dt.datetime.fromisoformat(iso)
        return d.strftime("%d %b %Y, %H:%M UTC")
    except (ValueError, TypeError):
        return iso or "—"


def _csv_bytes() -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Timestamp (UTC)", "User ID", "User name", "Channel", "Question"])
    for r in read_questions():
        w.writerow([r.get("ts", ""), r.get("user", ""), r.get("user_name", ""),
                    r.get("channel_type", ""), r.get("question", "")])
    return buf.getvalue().encode("utf-8")


def _page_html() -> str:
    rows = read_questions()
    banner = ""
    if not VIEW_TOKEN:
        banner = ('<div class="warn">⚠️ This page is open to anyone with the link. '
                  'Set <code>LOG_VIEW_TOKEN</code> and open with '
                  '<code>?token=…</code> to protect it.</div>')
    tok = f"?token={html.escape(VIEW_TOKEN)}" if VIEW_TOKEN else ""
    body = []
    for r in rows:
        who = html.escape(r.get("user_name") or r.get("user") or "—")
        body.append(
            "<tr>"
            f"<td class=ts>{html.escape(_fmt_ts(r.get('ts', '')))}</td>"
            f"<td>{who}</td>"
            f"<td>{html.escape(r.get('channel_type') or '—')}</td>"
            f"<td>{html.escape(r.get('question', ''))}</td>"
            "</tr>"
        )
    rows_html = "\n".join(body) or '<tr><td colspan=4 class=empty>No questions logged yet.</td></tr>'
    return f"""<!doctype html>
<html lang=en><head><meta charset=utf-8>
<meta name=viewport content="width=device-width, initial-scale=1">
<title>IT Helper — questions asked</title>
<style>
 body{{font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:0;background:#f7f7f8;color:#1d1d1f}}
 header{{background:#2c2d30;color:#fff;padding:16px 20px}}
 header h1{{margin:0;font-size:18px}} header .sub{{opacity:.75;font-size:13px;margin-top:4px}}
 .wrap{{padding:16px 20px}}
 .warn{{background:#fff4e5;border:1px solid #ffd699;color:#7a4f00;padding:8px 12px;border-radius:6px;margin-bottom:12px;font-size:13px}}
 .bar{{display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap}}
 input[type=search]{{flex:1;min-width:180px;padding:8px 10px;border:1px solid #ccc;border-radius:6px;font-size:14px}}
 a.btn{{background:#2c2d30;color:#fff;text-decoration:none;padding:8px 12px;border-radius:6px;font-size:13px}}
 .count{{color:#666;font-size:13px}}
 table{{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.08)}}
 th,td{{text-align:left;padding:10px 12px;border-bottom:1px solid #eee;font-size:14px;vertical-align:top}}
 th{{background:#fafafa;font-size:12px;text-transform:uppercase;letter-spacing:.03em;color:#666}}
 td.ts{{white-space:nowrap;color:#555;font-size:13px}}
 td.empty{{text-align:center;color:#999;padding:24px}}
 tr:hover td{{background:#fafbff}}
</style></head>
<body>
<header><h1>🛟 IT Helper — questions asked</h1>
<div class=sub>Every question sent to the bot, newest first.</div></header>
<div class=wrap>
{banner}
<div class=bar>
 <input type=search id=q placeholder="Filter questions, people…" oninput="flt()">
 <span class=count id=count>{len(rows)} question(s)</span>
 <a class=btn href="/questions.csv{tok}">⬇ Download CSV</a>
</div>
<table id=t><thead><tr><th>Time</th><th>Asked by</th><th>Where</th><th>Question</th></tr></thead>
<tbody>
{rows_html}
</tbody></table>
</div>
<script>
function flt(){{
 var v=document.getElementById('q').value.toLowerCase();
 var n=0, rows=document.querySelectorAll('#t tbody tr');
 rows.forEach(function(tr){{
   if(tr.querySelector('.empty')) return;
   var hit=tr.innerText.toLowerCase().indexOf(v)>=0;
   tr.style.display=hit?'':'none'; if(hit)n++;
 }});
 document.getElementById('count').innerText=n+' question(s)';
}}
</script>
</body></html>"""


class _Handler(BaseHTTPRequestHandler):
    server_version = "ITHelperLog/1.0"

    def _authorized(self) -> bool:
        if not VIEW_TOKEN:
            return True
        q = parse_qs(urlparse(self.path).query)
        return q.get("token", [""])[0] == VIEW_TOKEN

    def _send(self, code: int, ctype: str, body, download: str | None = None) -> None:
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        if download:
            self.send_header("Content-Disposition", f'attachment; filename="{download}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def do_GET(self):  # noqa: N802
        path = urlparse(self.path).path
        if path in ("/health", "/healthz"):
            return self._send(200, "text/plain; charset=utf-8", "ok")
        if not self._authorized():
            return self._send(401, "text/plain; charset=utf-8",
                              "Unauthorized — append ?token=<LOG_VIEW_TOKEN> to the URL.")
        if path == "/questions.csv":
            return self._send(200, "text/csv; charset=utf-8", _csv_bytes(),
                              download="question-log.csv")
        if path in ("/", "/index.html"):
            return self._send(200, "text/html; charset=utf-8", _page_html())
        return self._send(404, "text/plain; charset=utf-8", "Not found")

    do_HEAD = do_GET

    def log_message(self, *args):  # silence default per-request stderr logging
        return


def start_web_server() -> int:
    """Start the log web server in a daemon thread. Returns the bound port."""
    port = int(os.environ.get("PORT", "3000"))
    httpd = ThreadingHTTPServer(("0.0.0.0", port), _Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True,
                     name="qlog-web").start()
    if not VIEW_TOKEN:
        log.warning("LOG_VIEW_TOKEN not set — the question-log web page is OPEN to "
                    "anyone with the URL. Set LOG_VIEW_TOKEN to protect it.")
    log.info("Question-log web page on :%d (log file: %s)", port, LOG_PATH)
    return port
