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
- **"IT Issues"** _(optional, not present yet)_ — IT helpdesk/ticket log. Columns: Date Raised, Issue, Raised By, Priority, Status, Owner. Read automatically when added; until then the report shows a placeholder. The eventual feed (ticketing tool + email + Slack) is a separate integration task.

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

> **Note:** Reports are produced deterministically by `scripts/generate-report.py`
> (no LLM call at runtime). This file documents the intended structure; edit the
> script to change the actual output. The monthly report is the full weekly report
> plus the monthly deep-dive sections below. Ordered for an at-a-glance read.

Produce **two outputs**:

### Output 1: Slack Summary (`output/slack-summary.md`)

A concise monthly Slack post for the IT head, in this order:

1. **At a Glance** — overall status + the four health-check lights.
2. **Action Items** — prioritised IT-manager to-do list (most urgent first).
3. **Monthly Highlights** — fleet totals, assignments/replacements this month, aging counts.
4. **Stock Ready** — laptops ready + backup.
5. **Joiners** — next 7 / 30 days and 90-day forecast.
6. **IT Issues & Status** — open/resolved from the optional "IT Issues" sheet, or a placeholder.
7. **Monthly Spend** — app/subscription spend (this vs last month) and laptop spend vs budget, in USD.
8. **Procurement Recommendation** — order quantity from joiners + critical replacements vs stock.
9. **Renewals** — subscriptions due in the next 30 days.

### Output 2: Full Report (`output/full-report.md`)

All **weekly full-report sections** (At a Glance, Health Check, Action Items, Stock Ready,
Joiners Next Week, Laptop Aging with Remarks, IT Issues & Status, Spend MTD, Fleet Summary),
followed by these monthly deep-dive sections:

#### 📈 Monthly App Spend
- This month vs last month (USD) and month-over-month delta
- Top 5 app subscriptions by cost
- App spend by department

#### 💻 Monthly Laptop Spend
- Laptop spend this month (USD), joiners served, monthly budget and % used
- Laptops by model

#### 🛒 Procurement Recommendation
- Stock ready vs next-30-day demand (joiners + critical replacements) and 90-day joiner forecast
- Specific order quantity recommendation with a small buffer

#### 🧾 Budget vs Actual
- Planned procurement by department/model from procurement_plan "Laptop procurement plan" (INR → USD)

#### ✅ Onboarding Checklist Status
- Per recent joiner, which IT tasks are done vs pending from "Joining checklist"

## Rules

- Use today's date for all calculations.
- Compare with previous month data if available in the spreadsheets; otherwise note "prior month data not available."
- Format currency as USD ($) with comma separators. Laptop procurement & budget figures are recorded in INR in the source sheets and converted to USD at the `INR_TO_USD_RATE` rate (noted in the report footer).
- Sort all aging tables by age descending.
- Procurement recommendations should be actionable with specific quantities.
- Do NOT fabricate data — only report what exists in the Excel files. IT issues are read from an optional "IT Issues" sheet; when absent, show a placeholder rather than inventing tickets.
- If data is insufficient to calculate a metric, state "Insufficient data" rather than guessing.

## Output

Save the two files:
- `output/slack-summary.md`
- `output/full-report.md`
