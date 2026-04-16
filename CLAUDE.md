# IT Report Agent

This repo generates automated weekly and monthly IT asset & spend reports for Saras Solutions.

## How It Works

1. GitHub Actions fetches Excel files from SharePoint (asset inventory + spend tracker)
2. Claude Code reads the Excel data and generates reports using prompts in `prompts/`
3. A summary is posted to Slack `#it-reports`
4. The full report is saved as a ClickUp Doc in the IT team workspace

## Key Directories

- `data/` — downloaded Excel files (gitignored, populated at runtime)
- `output/` — generated reports (gitignored, populated at runtime)
- `prompts/` — Claude prompt templates for weekly and monthly reports
- `scripts/` — Python helper scripts (fetch, validate, post, create doc)

## Report Types

- **Weekly** (every Monday 9 AM IST): stock levels, assignments, replacements, aging alerts, spend snapshot
- **Monthly** (1st of month 9 AM IST): everything weekly + procurement recommendations, full aging analysis, spend trends, renewal calendar

## Rules for Claude

- Read Excel files using the xlsx skill
- Currency is INR (₹) with comma separators
- Assets older than 3.5 years (1,277 days) are flagged for replacement
- Never fabricate data — only report what exists in the spreadsheets
- Save output to `output/slack-summary.md` and `output/full-report.md`
