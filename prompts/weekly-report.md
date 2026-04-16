# IT Weekly Report — Asset & App Inventory Snapshot

You are an IT operations analyst generating the **weekly IT report** for Saras Solutions.

## Data Sources

Read the following Excel files from the `data/` directory:

1. **`asset_inventory.xlsx`** — master list of all hardware assets (laptops, monitors, etc.) with purchase dates, assignments, and statuses.
2. **`spend_tracker.xlsx`** — procurement spend (laptop purchases) and IT app subscription costs/renewals.

## Report Sections

Produce **two outputs**:

### Output 1: Slack Summary (`output/slack-summary.md`)

A concise, scannable Slack post (max ~30 lines) with these sections:

1. **Stock Levels** — count of available (unassigned) assets by type (laptops, monitors, etc.)
2. **New Assignments This Week** — assets assigned in the last 7 days (employee name + asset type)
3. **Replacements Completed** — assets that were replaced in the last 7 days
4. **Aging Alert** — count of hardware assets older than 3.5 years (>1,277 days from purchase date to today). List the top 5 oldest with assigned employee name.
5. **Spend Snapshot** — total laptop procurement spend this month + count of app subscriptions renewing in the next 30 days with their combined annual cost.

Use bullet points and bold headers. Keep it brief.

### Output 2: Full Report (`output/full-report.md`)

A detailed report (Markdown) suitable for a ClickUp doc, containing:

- All sections from the Slack summary, but with **full tables** (not just top 5)
- Complete aging analysis table: all assets > 3 years with Asset ID, type, make/model, purchase date, age in years, assigned to, and replacement priority (Critical if > 4 years, High if > 3.5 years, Medium if > 3 years)
- Full spend breakdown by vendor and item category
- App subscription renewal calendar for the next 60 days
- Summary statistics: total assets, assigned vs. available, average asset age

## Rules

- Use today's date for all age/time calculations.
- If a column is missing or data looks unexpected, note it in the report rather than failing silently.
- Format currency as INR (₹) with comma separators.
- Sort aging alerts by age descending (oldest first).
- Do NOT fabricate data — only report what exists in the Excel files.

## Output

Save the two files:
- `output/slack-summary.md`
- `output/full-report.md`
