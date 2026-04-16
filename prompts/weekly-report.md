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

Produce **two outputs**:

### Output 1: Slack Summary (`output/slack-summary.md`)

A concise, scannable Slack post (max ~30 lines) with these sections:

1. **Stock Levels** — count from "Laptop in stock" sheet + peripheral stock from "Other Assets Instock". Also count backup laptops from "Backup Laptops 3years old".
2. **New Assignments This Week** — from "Assset History" sheet, filter rows where "Assigned Date" is within the last 7 days. Show Username + Laptop Make/Model + whether New Joiner or Replacement.
3. **Replacements Completed** — from "Assset History" where "New Joiner/Replacement" = "Replacement" in the last 7 days, cross-referenced with "Laptop Returned".
4. **Aging Alert** — from "Laptop Assigned" sheet, calculate age using "Warranty Start Date" as purchase proxy (>3.5 years = flagged). Also check the "Laptop age" column if populated. List top 5 oldest with Employee Name.
5. **Spend Snapshot** — from spend_tracker "Sheet1", sum the current month's column for total app spend. Count subscriptions where "Renewal data" falls within the next 30 days.
6. **Upcoming Joiners** — from joiners_info "Joinings", list joiners with "Confirm DOJ" in the next 14 days (name, department, designation). Flag if laptop stock is insufficient.

Use bullet points and bold headers. Keep it brief.

### Output 2: Full Report (`output/full-report.md`)

A detailed report (Markdown) suitable for a ClickUp doc, containing:

- All sections from the Slack summary, but with **full tables** (not just top 5)
- Complete aging analysis table: all assets > 3 years with Asset ID, type, make/model, purchase date, age in years, assigned to, and replacement priority (Critical if > 4 years, High if > 3.5 years, Medium if > 3 years)
- Full spend breakdown by vendor and item category
- App subscription renewal calendar for the next 60 days
- Summary statistics: total assets, assigned vs. available, average asset age
- Upcoming joiners table (next 30 days) from joiners_info with department breakdown
- Procurement plan vs. actual comparison from procurement_plan "Actual Spends" sheet
- Onboarding checklist completion status from joiners_info "Joining checklist" (count of incomplete items per recent joiner)

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
