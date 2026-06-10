# IT Report Agent

This repo generates automated weekly and monthly IT asset & spend reports for Saras Solutions.

## How It Works

1. GitHub Actions fetches Excel files from SharePoint (asset inventory, spend tracker, procurement plan, joiner info)
2. Claude Code reads the Excel data and generates reports using prompts in `prompts/`
3. A summary is posted to Slack `#it-reports`
4. The full report is saved as a ClickUp Doc in the IT team workspace

## Key Directories

- `data/` — downloaded Excel files (gitignored, populated at runtime)
- `output/` — generated reports (gitignored, populated at runtime)
- `prompts/` — Claude prompt templates for weekly and monthly reports
- `scripts/` — Python helper scripts (fetch, validate, post, create doc)

## Report Types

- **Weekly** (every Monday 9 AM IST): stock levels, assignments, replacements, aging alerts, spend snapshot, upcoming joiners
- **Monthly** (1st of month 9 AM IST): everything weekly + procurement recommendations, full aging analysis, spend trends, renewal calendar, joiner onboarding status, budget vs. actual comparison

## Rules for Claude

- Read Excel files using the xlsx skill
- Currency is USD ($) with comma separators. Laptop procurement & budget figures are stored in INR in the source sheets and converted to USD via the `INR_TO_USD_RATE` env var (default ≈ ₹85.5/$); the rate is noted in the report footer.
- Assets older than 3.5 years (1,277 days) are flagged for replacement
- Never fabricate data — only report what exists in the spreadsheets. IT issues are pulled from the ClickUp IT ticket list by `scripts/fetch-issues.py` (set `CLICKUP_API_TOKEN`; list defaults to `CLICKUP_IT_ISSUES_LIST_ID`) into `data/it_issues.xlsx`; when no source is connected the report shows a placeholder. (Email + Slack issue feeds are still a future integration.)
- Save output to `output/slack-summary.md` and `output/full-report.md`
