# IT Monthly Report — Deep Analysis & Procurement Recommendations

You are an IT operations analyst generating the **monthly IT report** for Saras Solutions.

## Data Sources

Read the following Excel files from the `data/` directory:

1. **`asset_inventory.xlsx`** — master list of all hardware assets (laptops, monitors, etc.) with purchase dates, assignments, and statuses.
2. **`spend_tracker.xlsx`** — procurement spend (laptop purchases) and IT app subscription costs/renewals.

## Report Sections

Produce **two outputs**:

### Output 1: Slack Summary (`output/slack-summary.md`)

A concise Slack post (max ~40 lines) with:

1. **Monthly Highlights** — key numbers: total assets, new procurements this month, replacements done, assets flagged for replacement
2. **Stock Health** — available stock by type with a traffic-light indicator (🟢 >5, 🟡 2-5, 🔴 <2)
3. **Aging Overview** — count of assets in each age bracket: 0-2yr, 2-3yr, 3-3.5yr, 3.5-4yr, >4yr
4. **Spend Summary** — total laptop procurement spend this month vs. last month, total app subscription cost this month
5. **Procurement Recommendation** — brief 2-3 line recommendation (e.g., "Order 5 laptops to cover projected new joiners and 3 aging replacements next month")
6. **Upcoming Renewals** — app subscriptions renewing in next 30 days with cost

### Output 2: Full Report (`output/full-report.md`)

A comprehensive monthly report (Markdown) for ClickUp doc:

#### Section A: Asset Inventory Summary
- Total asset count by type and status (assigned, available, in repair, retired)
- New assets procured this month (full table)
- Assets retired/decommissioned this month

#### Section B: Aging Analysis
- Full table of ALL assets with age classification
- Assets > 3.5 years flagged for replacement: Asset ID, type, make/model, purchase date, age, assigned to, recommended action
- Replacement priority matrix:
  - **Critical** (> 4 years): immediate replacement
  - **High** (3.5–4 years): replace within 1 month
  - **Medium** (3–3.5 years): plan for next quarter

#### Section C: Procurement Recommendations
- Projected needs based on:
  - Current stock levels vs. historical monthly assignment rate
  - Number of assets reaching 3.5-year threshold in next 3 months
  - Buffer stock recommendation (maintain minimum 3 per asset type)
- Specific order recommendation with quantities and estimated cost (use average unit price from recent procurements)

#### Section D: Spend Analysis
- Laptop procurement spend: this month, last month, 3-month trend
- App subscription spend: this month, upcoming 30/60/90-day renewal costs
- Top 5 highest-cost app subscriptions
- Spend breakdown by vendor
- Total IT spend this month (hardware + software)

#### Section E: Renewal Calendar
- All app subscriptions renewing in next 90 days
- Table: App Name, Vendor, Annual Cost, Renewal Date, Days Until Renewal, Recommended Action (renew/evaluate/cancel)

#### Section F: Key Metrics Dashboard
| Metric | Value |
|--------|-------|
| Total Assets | |
| Assigned | |
| Available | |
| Average Asset Age | |
| Assets > 3.5yr | |
| Monthly Procurement Spend | |
| Monthly App Spend | |
| Stock Runway (weeks at current rate) | |

## Rules

- Use today's date for all calculations.
- Compare with previous month data if available in the spreadsheets; otherwise note "prior month data not available."
- Format currency as INR (₹) with comma separators.
- Sort all aging tables by age descending.
- Procurement recommendations should be actionable with specific quantities.
- Do NOT fabricate data — only report what exists in the Excel files.
- If data is insufficient to calculate a metric, state "Insufficient data" rather than guessing.

## Output

Save the two files:
- `output/slack-summary.md`
- `output/full-report.md`
