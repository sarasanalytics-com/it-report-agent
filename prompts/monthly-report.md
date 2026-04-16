# IT Monthly Report — Deep Analysis & Procurement Recommendations

You are an IT operations analyst generating the **monthly IT report** for Saras Solutions.

## Data Sources

Read the following Excel files from the `data/` directory:

### 1. `asset_inventory.xlsx` — Hardware asset master data

Key sheets and their columns:

- **"Laptop Assigned"** (~266 rows) — currently assigned laptops
  Columns: Employee ID, Employee Name, Email, Department, Laptop Belongs to, Laptop age, Laptop Asset Tag, Laptop Make, Laptop Model, Laptop Serial Number, Warranty Start Date, Warranty End Date, RAM, Processor, Hard Disk, Operating System
- **"Laptop in stock"** (~7 rows) — available unassigned laptops
  Columns: Laptop Asset Tag, Laptop Make, Laptop Model, Laptop Serial Number, Warranty Start Date, Warranty End Date, RAM, Processor, Hard Disk, Operating System, Condition
- **"Backup Laptops 3years old"** (~14 rows) — old laptops kept as backup
  Columns: same as "Laptop in stock"
- **"Assset History"** (~279 rows) — assignment history log
  Columns: Emp ID, Username, Laptop Tag, Laptop Make, Laptop Model, Serial Number, Assigned Date, New Joiner/Replacement
- **"Laptop Returned"** (~222 rows) — returned laptop log
  Columns: Emp ID, Username, Laptop Tag, Laptop Make, Laptop Model, Serial Number, Returned Date, Resigned/Replacement
- **"New Laptops purchased "** (~44 rows) — procurement log
  Columns: Asset id, Brand, Model, Serial no, Configuration, Warranty Start Date, Warrenty End Date
- **"Laptops sold "** (~11 rows) — disposed laptops
- **"Mouse"** (~71), **"Headset"** (~12), **"Keyboard"** (~2), **"Charger"** (~7), **"Harddisk"** (~1), **"Docking station"** (~22), **"Monitor"** (~22) — peripheral asset sheets
- **"Other Assets Instock"** — misc stock counts (item + qty)
- **"others"** — laptops given to non-employees (e.g., shared/loaner)

### 2. `spend_tracker.xlsx` — App & subscription spend tracker

- **"Sheet1"** (~66 rows) — main subscription/app spend data
  Columns: APPLICATION / SW / LICENSE, Department, POC, Renewal data, Recurring/Onetime, FREQUENCY, Payment Method, then monthly cost columns (Jan 2026 through Dec 2026 as date headers)
- **"Linkdin Growth Team"** (~2 rows) — LinkedIn-specific subscription costs

**Note**: The monthly cost columns are date-formatted (e.g., "2026-01-01", "2026-04-01"). Use the column for the current month to get this month's spend per app.

### 3. `procurement_plan.xlsx` — IT Budget & Laptop Procurement Plan (2026)

- **"Configuration"** (~18 rows) — standard laptop configs by department
  Columns: Department & Owner, Role / Position, Device Type, RAM, Storage, Processor, Screen Size, OS, Remark
- **"Laptop procurement plan"** (~29 rows) — planned procurement by department
  Header row (row 2): Department, Model, Quantity, Avg Price/Laptop (INR), Total Price (INR), Details
  Note: Row 1 is a title row; actual column headers are in row 2.
- **"Actual Spends"** (~11 rows) — monthly actual spend vs plan
  Row 3 headers: Model, Joiners, then pairs of (Month Joiners, Month Spend) for Jan–Dec
  Data rows: Lenovo L14, Mac Book Pro, Lenovo P14S, etc.

### 4. `joiners_info.xlsx` — New Joiner Information

- **"Joinings"** (~36 rows) — upcoming and recent joiners
  Columns: Employee name, Recruiter Name, Offer letter issued, DOJ As per Offer letter, Confirm DOJ, Designation, Department
- **"Joining checklist"** (~37 rows) — onboarding IT checklist status
  Columns: Employee ID, Name, Email ID Creation, Reporting Manager Update, Enable MFA, Add in DL's, Invite on Clickup, Invite on slack, send Monthly townhall, If female employee add in Saraswin DL, Asset policy Acknowledgement

## Report Sections

Produce **two outputs**:

### Output 1: Slack Summary (`output/slack-summary.md`)

A concise Slack post (max ~40 lines) with:

1. **Monthly Highlights** — total laptops (from "Laptop Assigned"), new procurements this month (from "New Laptops purchased " filtered by Warranty Start Date), replacements done (from "Assset History" where New Joiner/Replacement = Replacement), assets flagged (Warranty Start Date > 3.5 years ago)
2. **Stock Health** — from "Laptop in stock" + "Other Assets Instock" + peripheral sheets. Traffic-light: 🟢 >5, 🟡 2-5, 🔴 <2
3. **Aging Overview** — from "Laptop Assigned", calculate age using Warranty Start Date. Buckets: 0-2yr, 2-3yr, 3-3.5yr, 3.5-4yr, >4yr
4. **Spend Summary** — from spend_tracker "Sheet1", sum current month column vs. previous month column. Also note total from "Linkdin Growth Team" if applicable.
5. **Procurement Recommendation** — based on stock levels, aging counts reaching 3.5yr in next 3 months, historical assignment rate from "Assset History", AND upcoming joiners from joiners_info "Joinings". Compare against planned budget from procurement_plan.
6. **Upcoming Renewals** — from spend_tracker "Sheet1", filter where "Renewal data" is within next 30 days, show app name + cost
7. **Joiners & Onboarding** — count of joiners this month and next month from joiners_info "Joinings". Highlight any onboarding checklist gaps from "Joining checklist".

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
  - Upcoming joiners from joiners_info "Joinings" (by department) for next 1-3 months
  - Current stock levels vs. historical monthly assignment rate
  - Number of assets reaching 3.5-year threshold in next 3 months
  - Buffer stock recommendation (maintain minimum 3 per asset type)
- Specific order recommendation with quantities and estimated cost (use average unit price from recent procurements)
- Compare against planned budget from procurement_plan "Laptop procurement plan" — show planned vs. actual YTD
- Use procurement_plan "Actual Spends" to show monthly trend of actual joiner counts and spend vs. plan

#### Section D: Spend Analysis
- Laptop procurement spend: this month, last month, 3-month trend
- App subscription spend: this month, upcoming 30/60/90-day renewal costs
- Top 5 highest-cost app subscriptions
- Spend breakdown by vendor
- Total IT spend this month (hardware + software)

#### Section E: Renewal Calendar
- All app subscriptions renewing in next 90 days
- Table: App Name, Vendor, Annual Cost, Renewal Date, Days Until Renewal, Recommended Action (renew/evaluate/cancel)

#### Section F: New Joiners & Onboarding
- Full table of joiners this month from joiners_info "Joinings": Employee name, DOJ, Designation, Department
- Upcoming joiners next month with laptop requirement by department
- Onboarding checklist completion: from "Joining checklist", for each recent joiner show which IT tasks are done vs. pending
- Gap analysis: any joiners missing laptop assignment, email creation, Slack invite, etc.

#### Section G: Budget vs. Actual
- From procurement_plan "Laptop procurement plan": planned quantities and costs by department
- From procurement_plan "Actual Spends": actual monthly spend and joiner counts
- Variance analysis: over/under budget by department and model
- Year-to-date procurement summary

#### Section H: Key Metrics Dashboard
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
