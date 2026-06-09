# IT Weekly Report — Asset & App Inventory Snapshot

You are an IT operations analyst generating the **weekly IT report** for Saras Solutions.

## Data Sources

Read the following Excel files from the `data/` directory:

### 1. `asset_inventory.xlsx` — Hardware asset master data

Key sheets and their columns:

- **"Laptop Assigned"** — currently assigned laptops
  Columns: Employee ID, Employee Name, Email, Department, Laptop Belongs to, Laptop age, Laptop Asset Tag, Laptop Make, Laptop Model, Laptop Serial Number, Warranty Start Date, Warranty End Date, RAM, Processor, Hard Disk, Operating System
- **"Laptop in stock"** — available unassigned laptops
  Columns: Laptop Asset Tag, Laptop Make, Laptop Model, Laptop Serial Number, Warranty Start Date, Warranty End Date, RAM, Processor, Hard Disk, Operating System, Condition
- **"Backup Laptops 3years old"** — old laptops kept as backup
  Columns: same as "Laptop in stock"
- **"Assset History"** — assignment history log
  Columns: Emp ID, Username, Laptop Tag, Laptop Make, Laptop Model, Serial Number, Assigned Date, New Joiner/Replacement
- **"Laptop Returned"** — returned laptop log
  Columns: Emp ID, Username, Laptop Tag, Laptop Make, Laptop Model, Serial Number, Returned Date, Resigned/Replacement
- **"New Laptops purchased "** — procurement log
  Columns: Asset id, Brand, Model, Serial no, Configuration, Warranty Start Date, Warrenty End Date
- **"Laptops sold "** — disposed laptops
- **"Mouse"**, **"Headset"**, **"Keyboard"**, **"Charger"**, **"Harddisk"**, **"Docking station"**, **"Monitor"** — peripheral asset sheets
- **"Other Assets Instock"** — misc stock counts
- **"IT Issues"** _(optional, not present yet)_ — IT helpdesk/ticket log. Columns: Date Raised, Issue, Raised By, Priority, Status, Owner. Read automatically when added; until then the report shows a placeholder. The eventual feed (ticketing tool + email + Slack) is a separate integration task.

### 2. `spend_tracker.xlsx` — App & subscription spend tracker

- **"Sheet1"** — main subscription/app spend data
  Columns: APPLICATION / SW / LICENSE, Department, POC, Renewal data, Recurring/Onetime, FREQUENCY, Payment Method, then monthly cost columns (Jan 2026 through Dec 2026)
- **"Linkdin Growth Team"** — LinkedIn-specific subscription costs

### 3. `procurement_plan.xlsx` — IT Budget & Laptop Procurement Plan (2026)

- **"Configuration"** — standard laptop configs by department
  Columns: Department & Owner, Role / Position, Device Type, RAM, Storage, Processor, Screen Size, OS, Remark
- **"Laptop procurement plan"** — planned procurement by department
  Header row (row 2): Department, Model, Quantity, Avg Price/Laptop (INR), Total Price (INR), Details
  Note: Row 1 is a title row; actual column headers are in row 2.
- **"Actual Spends"** — monthly actual spend vs plan
  Row 3 headers: Model, Joiners, then pairs of (Month Joiners, Month Spend) for Jan–Dec
  Data rows: Lenovo L14, Mac Book Pro, Lenovo P14S, etc.

### 4. `joiners_info.xlsx` — New Joiner Information

- **"Joinings"** (~36 rows) — upcoming and recent joiners
  Columns: Employee name, Recruiter Name, Offer letter issued, DOJ As per Offer letter, Confirm DOJ, Designation, Department
- **"Joining checklist"** (~37 rows) — onboarding IT checklist
  Columns: Employee ID, Name, Email ID Creation, Reporting Manager Update, Enable MFA, Add in DL's, Invite on Clickup, Invite on slack, send Monthly townhall, If female employee add in Saraswin DL, Asset policy Acknowledgement

## Report Sections

> **Note:** Reports are produced deterministically by `scripts/generate-report.py`
> (no LLM call at runtime). This file documents the intended structure; edit the
> script to change the actual output. Sections are ordered so an IT head can read
> the situation at a glance, top-down.

Produce **two outputs**:

### Output 1: Slack Summary (`output/slack-summary.md`)

A concise, scannable Slack post for the IT head, in this order:

1. **At a Glance** — overall status (🟢/🟡/🔴) + the four health-check lights (Stock · Aging · Joiner Prep · Spend).
2. **Action Items** — the prioritised to-do list for the IT manager, synthesised from stock gaps, aging, onboarding gaps, renewals and open IT issues. Most urgent first.
3. **Stock Ready** — laptops ready to assign (+ backup count), with week-over-week delta.
4. **Joiners Next Week** — joiners with "Confirm DOJ" in the next 7 days (name, department, DOJ, required laptop config).
5. **Laptop Aging** — count over 3.5 years (and critical count), with the top critical laptops and an auto-generated remark each.
6. **IT Issues & Status** — open vs resolved counts from the optional "IT Issues" sheet; a placeholder note when no source is connected.
7. **Spend MTD** — app/subscription spend and laptop spend (USD), plus renewals due in 30 days.

Use bullet points and bold headers. Keep it brief.

### Output 2: Full Report (`output/full-report.md`)

A detailed Markdown report for a ClickUp doc / Word export, same section order with full tables:

- **At a Glance** — headline KPI table + the single top action.
- **Health Check** — traffic-light table.
- **Action Items for IT Manager** — numbered, prioritised table.
- **Stock Ready** — stock levels vs last week, other assets in stock, activity this week vs last.
- **Joiners Next Week** — full table + stock-vs-joiners analysis + upcoming joiners (next 30 days).
- **Laptop Aging** — all assets > 3.5 years with employee, department, asset tag, make/model, purchase date, age, priority (Critical > 4yr, High 3.5–4yr) and an auto-generated **Remarks** column (replacement urgency + warranty status). Plus an age-distribution table.
- **IT Issues & Status** — table from the optional "IT Issues" sheet, or a placeholder.
- **Spend (MTD)** — app/subscription categories + renewal calendar + laptop procurement summary and purchases (all USD).
- **Fleet Summary** — totals, assigned vs available, average age.

## Rules

- Use today's date for all age/time calculations.
- If a column is missing or data looks unexpected, note it in the report rather than failing silently.
- Format currency as USD ($) with comma separators. Laptop procurement & budget figures are recorded in INR in the source sheets and converted to USD at the `INR_TO_USD_RATE` rate (noted in the report footer).
- Sort aging alerts by age descending (oldest first).
- Do NOT fabricate data — only report what exists in the Excel files. IT issues are read from an optional "IT Issues" sheet; when absent, show a placeholder rather than inventing tickets.

## Output

Save the two files:
- `output/slack-summary.md`
- `output/full-report.md`
