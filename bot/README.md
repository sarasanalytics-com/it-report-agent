# IT Helper — Slack Q&A bot

An always-on Slack bot that lets the **Head of HR** (non-technical) ask, in plain
language, about company laptops, new-joiner readiness, open IT requests and
spend — and get instant, grounded answers.

It reuses this repo's existing pipeline to pull **live data** (SharePoint
spreadsheets + ClickUp tickets), turns it into a report, and asks **Claude** to
answer the question using only that data (so it won't make things up).

```
Slack (@mention or DM)
        │
        ▼
  bot/app.py  ──Socket Mode──►  Slack
        │
        ▼
 bot/answer.py ──► fetch-excel.py + fetch-issues.py + generate-report.py
        │              (cached, refreshed on a TTL)
        ▼
   Claude (Anthropic API)  ──►  plain-language answer
```

## What she can ask

- *Will the new joiners have laptops ready?*
- *Is there a laptop ready for <name>?*
- *How many spare laptops do we have right now?*
- *What IT requests are open at the moment, and who raised them?*
- *How much did we spend on software this month?*
- *Are there any vendor bills pending?*

The bot answers in simple, non-technical language and only from the latest data.

## One-time setup

### 1. Create the Slack app
1. Go to <https://api.slack.com/apps> → **Create New App** → **From an app manifest**.
2. Pick the workspace, paste in [`bot/manifest.yaml`](./manifest.yaml), create.
3. **Basic Information → App-Level Tokens →** *Generate Token and Scopes*: add
   `connections:write`. Copy the `xapp-…` token → `SLACK_APP_TOKEN`.
4. **Install App** (OAuth & Permissions) → install to workspace. Copy the
   **Bot User OAuth Token** `xoxb-…` → `SLACK_BOT_TOKEN`.
5. Invite the bot to a channel (`/invite @IT Helper`) and/or just DM it.

### 2. Configure env
```bash
cp bot/.env.example bot/.env
# fill in Slack + Anthropic + the same SharePoint/ClickUp secrets the
# report workflow uses
```

### 3. Run it

**Locally (to try it):**
```bash
pip install -r requirements.txt -r bot/requirements.txt
set -a; source bot/.env; set +a
python bot/app.py
```

**Test just the answer engine (no Slack needed):**
```bash
python bot/answer.py "will the new joiners have laptops ready?"
python bot/answer.py --no-refresh "how many spare laptops do we have?"   # use last data
```

**With Docker:**
```bash
docker build -f bot/Dockerfile -t it-bot .
docker run --env-file bot/.env it-bot
```

## Deploy (always-on, cloud)

Socket Mode means **no public URL** — any host that can run a long-lived process
works. Recommended easy options:

- **Render** → New **Background Worker** → connect this repo → Docker (or
  `pip install -r requirements.txt -r bot/requirements.txt` + start command
  `python bot/app.py`) → add the env vars from `bot/.env`.
- **Railway / Fly.io** → deploy the Dockerfile → set the same env vars.

Pick the smallest instance; the bot is idle until asked.

## Notes & guardrails

- **Grounded answers only.** The system prompt restricts Claude to the fetched
  data; if something isn't there it says so rather than guessing.
- **Freshness.** Data is cached for `DATA_REFRESH_TTL` seconds (default 10 min)
  and re-fetched on the next question after that.
- **Cost.** Defaults to `claude-sonnet-4-6`; set `ANTHROPIC_MODEL=claude-haiku-4-5`
  for the cheapest/fastest option.
- **Access control.** Set `BOT_ALLOWED_USERS` to a comma-separated list of Slack
  user IDs (e.g. the HR head + IT) — only they get answers; everyone else gets a
  polite decline. Find an ID in Slack: click the person's name → **More** →
  **Copy member ID** (looks like `U01ABCDEF`). If left empty the bot answers
  anyone who can reach it (it logs a warning on startup).
