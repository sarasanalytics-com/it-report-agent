#!/usr/bin/env python3
"""
Generate IT weekly or monthly reports from Excel data.

Usage:
    python generate-report.py weekly
    python generate-report.py monthly

Reads Excel files from data/ and writes:
    output/slack-summary.md
    output/full-report.md
"""

from __future__ import annotations

import sys
import pathlib
import datetime as dt
from datetime import timedelta
from collections import defaultdict
from typing import Optional

import openpyxl

DATA_DIR = pathlib.Path(__file__).resolve().parent.parent / "data"
OUTPUT_DIR = pathlib.Path(__file__).resolve().parent.parent / "output"
TODAY = dt.datetime.now().date()
AGE_THRESHOLD_DAYS = int(3.5 * 365)  # 1277 days


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_date(val) -> Optional[dt.date]:
    if val is None:
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return dt.datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def age_years(dt: datetime.date) -> float:
    return (TODAY - dt).days / 365.25


def fmt_usd(amount) -> str:
    try:
        n = float(amount)
    except (TypeError, ValueError):
        return "N/A"
    if n >= 1_000_000:
        return f"${n/1_000_000:,.2f}M"
    if n >= 1_000:
        return f"${n:,.2f}"
    return f"${n:,.2f}"


def read_sheet(wb, sheet_name: str, header_row: int = 1) -> list[dict]:
    """Read a sheet into a list of dicts. Returns [] if sheet missing."""
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=header_row, values_only=True))
    if len(rows) < 2:
        return []
    headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows[0])]
    result = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        result.append(dict(zip(headers, row)))
    return result


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_data() -> dict:
    asset_wb = openpyxl.load_workbook(DATA_DIR / "asset_inventory.xlsx", read_only=True, data_only=True)
    spend_wb = openpyxl.load_workbook(DATA_DIR / "spend_tracker.xlsx", read_only=True, data_only=True)
    proc_wb = openpyxl.load_workbook(DATA_DIR / "procurement_plan.xlsx", read_only=True, data_only=True)
    join_wb = openpyxl.load_workbook(DATA_DIR / "joiners_info.xlsx", read_only=True, data_only=True)

    data = {
        "assigned": read_sheet(asset_wb, "Laptop Assigned"),
        "in_stock": read_sheet(asset_wb, "Laptop in stock"),
        "backup": read_sheet(asset_wb, "Backup Laptops 3years old"),
        "history": read_sheet(asset_wb, "Assset History"),
        "returned": read_sheet(asset_wb, "Laptop Returned"),
        "purchased": read_sheet(asset_wb, "New Laptops purchased "),
        "sold": read_sheet(asset_wb, "Laptops sold "),
        "mouse": read_sheet(asset_wb, "Mouse"),
        "headset": read_sheet(asset_wb, "Headset"),
        "keyboard": read_sheet(asset_wb, "Keyboard"),
        "charger": read_sheet(asset_wb, "Charger"),
        "docking": read_sheet(asset_wb, "Docking station"),
        "monitor": read_sheet(asset_wb, "Monitor"),
        "other_stock": read_sheet(asset_wb, "Other Assets Instock"),
        "spend": read_sheet(spend_wb, "Sheet1"),
        "joinings": read_sheet(join_wb, "Joinings"),
        "checklist": read_sheet(join_wb, "Joining checklist"),
        "proc_plan": read_sheet(proc_wb, "Laptop procurement plan", header_row=2),
        "actual_spend": read_sheet(proc_wb, "Actual Spends", header_row=3),
    }

    for wb in (asset_wb, spend_wb, proc_wb, join_wb):
        wb.close()

    return data


# ---------------------------------------------------------------------------
# Analysis functions
# ---------------------------------------------------------------------------

def get_stock_summary(data: dict) -> dict:
    return {
        "Laptops (ready)": len(data["in_stock"]),
        "Laptops (3yr+ backup)": len(data["backup"]),
    }


def get_recent_assignments(data: dict, days: int = 7) -> list[dict]:
    cutoff = TODAY - timedelta(days=days)
    results = []
    for row in data["history"]:
        dt = parse_date(row.get("Assigned Date"))
        if dt and dt >= cutoff:
            results.append(row)
    return results


def get_recent_returns(data: dict, days: int = 7) -> list[dict]:
    cutoff = TODAY - timedelta(days=days)
    results = []
    for row in data["returned"]:
        dt = parse_date(row.get("Returned Date"))
        if dt and dt >= cutoff:
            results.append(row)
    return results


def get_aging_laptops(data: dict) -> list[dict]:
    """Return assigned laptops older than 3.5 years, sorted oldest first."""
    aging = []
    for row in data["assigned"]:
        dt = parse_date(row.get("Warranty Start Date"))
        if dt and (TODAY - dt).days > AGE_THRESHOLD_DAYS:
            aging.append({
                "employee": row.get("Employee Name", "Unknown"),
                "tag": row.get("Laptop Asset Tag", ""),
                "make": row.get("Laptop Make", ""),
                "model": row.get("Laptop Model", ""),
                "start_date": dt,
                "age_years": round(age_years(dt), 1),
                "priority": "Critical" if age_years(dt) > 4 else "High",
            })
    aging.sort(key=lambda x: x["age_years"], reverse=True)
    return aging


def get_age_distribution(data: dict) -> dict:
    buckets = {"0-2yr": 0, "2-3yr": 0, "3-3.5yr": 0, "3.5-4yr": 0, ">4yr": 0}
    for row in data["assigned"]:
        dt = parse_date(row.get("Warranty Start Date"))
        if not dt:
            continue
        yrs = age_years(dt)
        if yrs > 4:
            buckets[">4yr"] += 1
        elif yrs > 3.5:
            buckets["3.5-4yr"] += 1
        elif yrs > 3:
            buckets["3-3.5yr"] += 1
        elif yrs > 2:
            buckets["2-3yr"] += 1
        else:
            buckets["0-2yr"] += 1
    return buckets


def get_laptop_spend(data: dict) -> dict:
    """Extract laptop procurement spend from Actual Spends sheet.

    The sheet has columns like: Model, Joiners, Jan Joiners, Jan Spend,
    Feb Joiners, Feb Spend, March Joiners, Mar Spend, etc.
    Returns dict with keys: models (list of per-model data), total_joiners, total_spend.
    """
    MONTH_ABBREVS = {
        1: ["jan"], 2: ["feb"], 3: ["mar", "march"], 4: ["apr", "april"],
        5: ["may"], 6: ["jun", "june"], 7: ["jul", "july"], 8: ["aug"],
        9: ["sep", "sept"], 10: ["oct"], 11: ["nov"], 12: ["dec"],
    }
    abbrevs = MONTH_ABBREVS.get(TODAY.month, [])

    result = {"models": [], "total_joiners": 0, "total_spend": 0.0}
    for row in data["actual_spend"]:
        model = row.get("Model", "")
        if not model or str(model).strip().lower() in ("", "none", "total"):
            continue

        # Find joiners and spend columns for current month
        joiners = 0
        spend = 0.0
        for key, val in row.items():
            key_lower = str(key).strip().lower()
            for abbr in abbrevs:
                if abbr in key_lower and "joiner" in key_lower:
                    joiners = int(val) if val and str(val).strip() not in ("", "None") else 0
                elif abbr in key_lower and "spend" in key_lower:
                    spend = float(val) if val and str(val).strip() not in ("", "None") else 0.0

        if joiners or spend:
            result["models"].append({
                "model": str(model).strip(),
                "joiners": joiners,
                "spend": spend,
            })
            result["total_joiners"] += joiners
            result["total_spend"] += spend

    return result


def get_recent_purchases(data: dict, days: int = 30) -> list[dict]:
    """Get laptops purchased within the given number of days."""
    cutoff = TODAY - timedelta(days=days)
    purchases = []
    for row in data["purchased"]:
        d = parse_date(row.get("Warranty Start Date"))
        if d and d >= cutoff:
            purchases.append({
                "brand": row.get("Brand", ""),
                "model": row.get("Model", ""),
                "serial": row.get("Serial no", ""),
                "date": d,
            })
    purchases.sort(key=lambda x: x["date"], reverse=True)
    return purchases


# Row names in spend tracker that are laptop/hardware costs, not app subscriptions
HARDWARE_SPEND_KEYWORDS = ["laptop", "procurement", "antivirus", "mdm"]


def _is_hardware_row(row: dict) -> bool:
    app_name = str(row.get("APPLICATION / SW / LICENSE", "")).lower()
    return any(kw in app_name for kw in HARDWARE_SPEND_KEYWORDS)


def get_current_month_spend(data: dict) -> tuple[float, list[dict]]:
    """Get total app spend for current month and upcoming renewals.
    Excludes hardware/laptop rows (tracked separately in laptop procurement)."""
    # Find the column matching current month
    month_key = None
    for row in data["spend"]:
        for key in row:
            dt = parse_date(key)
            if dt and dt.year == TODAY.year and dt.month == TODAY.month:
                month_key = key
                break
        if month_key:
            break

    total = 0.0
    if month_key:
        for row in data["spend"]:
            if _is_hardware_row(row):
                continue
            val = row.get(month_key)
            if val and isinstance(val, (int, float)):
                total += float(val)

    # Upcoming renewals
    renewals = []
    cutoff = TODAY + timedelta(days=30)
    for row in data["spend"]:
        if _is_hardware_row(row):
            continue
        rd = parse_date(row.get("Renewal data"))
        if rd and TODAY <= rd <= cutoff:
            renewals.append({
                "app": row.get("APPLICATION / SW / LICENSE", "Unknown"),
                "date": rd,
                "dept": row.get("Department", ""),
                "frequency": row.get("FREQUENCY", ""),
            })
    renewals.sort(key=lambda x: x["date"])
    return total, renewals


def get_upcoming_joiners(data: dict, days: int = 14) -> list[dict]:
    cutoff = TODAY + timedelta(days=days)
    joiners = []
    for row in data["joinings"]:
        dt = parse_date(row.get("Confirm DOJ") or row.get("DOJ As per Offer letter"))
        if dt and TODAY <= dt <= cutoff:
            joiners.append({
                "name": row.get("Employee name", "Unknown"),
                "doj": dt,
                "designation": row.get("Designation", ""),
                "department": row.get("Department", ""),
            })
    joiners.sort(key=lambda x: x["doj"])
    return joiners


# ---------------------------------------------------------------------------
# Report generators
# ---------------------------------------------------------------------------

def generate_weekly_slack(data: dict) -> str:
    lines = [f"*📊 IT Weekly Report — {TODAY.strftime('%d %B %Y')}*\n"]

    # 1. Stock
    stock = get_stock_summary(data)
    lines.append("*1. Stock Levels*")
    for item, count in stock.items():
        icon = "🟢" if count > 5 else ("🟡" if count >= 2 else "🔴")
        lines.append(f"• {icon} {item}: {count}")

    # 2. New Assignments
    assignments = get_recent_assignments(data, 7)
    lines.append(f"\n*2. New Assignments This Week* ({len(assignments)})")
    for a in assignments[:5]:
        atype = a.get("New Joiner/Replacement", "")
        lines.append(f"• {a.get('Username', 'N/A')} — {a.get('Laptop Make', '')} {a.get('Laptop Model', '')} ({atype})")
    if len(assignments) > 5:
        lines.append(f"  _…and {len(assignments)-5} more_")

    # 3. Replacements
    replacements = [a for a in assignments if str(a.get("New Joiner/Replacement", "")).lower() == "replacement"]
    lines.append(f"\n*3. Replacements Completed* ({len(replacements)})")
    for r in replacements[:3]:
        lines.append(f"• {r.get('Username', 'N/A')} — {r.get('Laptop Make', '')} {r.get('Laptop Model', '')}")

    # 4. Aging
    aging = get_aging_laptops(data)
    lines.append(f"\n*4. Aging Alert* ({len(aging)} laptops > 3.5 years)")
    for a in aging[:5]:
        lines.append(f"• {a['employee']} — {a['make']} {a['model']} ({a['age_years']}yr, {a['priority']})")

    # 5. Laptop Procurement
    laptop_spend = get_laptop_spend(data)
    purchases = get_recent_purchases(data, 7)
    lines.append(f"\n*5. Laptop Procurement — {TODAY.strftime('%B %Y')}*")
    if laptop_spend["models"]:
        lines.append(f"• Joiners this month: {laptop_spend['total_joiners']}")
        lines.append(f"• Laptop spend this month: {fmt_usd(laptop_spend['total_spend'])}")
        for m in laptop_spend["models"]:
            lines.append(f"  - {m['model']}: {m['joiners']} joiners, {fmt_usd(m['spend'])}")
    else:
        lines.append("• No laptop procurement data for this month")
    if purchases:
        lines.append(f"• New laptops purchased this week: {len(purchases)}")
        for p in purchases[:3]:
            lines.append(f"  - {p['brand']} {p['model']} ({p['date'].strftime('%d %b')})")

    # 6. App Spend
    total_spend, renewals = get_current_month_spend(data)
    lines.append(f"\n*6. App Spend — {TODAY.strftime('%B %Y')}*")
    lines.append(f"• Total this month: {fmt_usd(total_spend)}")
    lines.append(f"• Renewals in next 30 days: {len(renewals)}")
    for r in renewals[:3]:
        lines.append(f"  - {r['app']} ({r['date'].strftime('%d %b')})")

    # 7. Upcoming Joiners
    joiners = get_upcoming_joiners(data, 14)
    lines.append(f"\n*7. Upcoming Joiners (next 14 days)* ({len(joiners)})")
    for j in joiners[:5]:
        lines.append(f"• {j['name']} — {j['department']}, {j['designation']} (DOJ: {j['doj'].strftime('%d %b')})")
    if not joiners:
        lines.append("• None in the next 14 days")

    stock_laptops = stock["Laptops (ready)"]
    if joiners and stock_laptops < len(joiners):
        lines.append(f"\n⚠️ *Stock alert*: {stock_laptops} laptops available but {len(joiners)} joiners expected!")

    lines.append(f"\n_Generated: {TODAY.strftime('%d %B %Y')}_")
    return "\n".join(lines)


def generate_weekly_full(data: dict) -> str:
    lines = [f"# IT Weekly Report — {TODAY.strftime('%d %B %Y')}\n"]

    # Stock
    stock = get_stock_summary(data)
    lines.append("## 1. Stock Levels\n")
    lines.append("| Asset Type | Available |")
    lines.append("|------------|-----------|")
    for item, count in stock.items():
        lines.append(f"| {item} | {count} |")

    # Other assets in stock
    if data["other_stock"]:
        lines.append("\n### Other Assets in Stock\n")
        lines.append("| Item | Qty |")
        lines.append("|------|-----|")
        for row in data["other_stock"]:
            item = row.get("Other Assets Instock", "")
            qty = row.get("Qty", "")
            if item:
                lines.append(f"| {item} | {qty} |")

    # Assignments
    assignments = get_recent_assignments(data, 7)
    lines.append(f"\n## 2. New Assignments This Week ({len(assignments)})\n")
    if assignments:
        lines.append("| Employee | Laptop | Type |")
        lines.append("|----------|--------|------|")
        for a in assignments:
            lines.append(f"| {a.get('Username', 'N/A')} | {a.get('Laptop Make', '')} {a.get('Laptop Model', '')} | {a.get('New Joiner/Replacement', '')} |")
    else:
        lines.append("No new assignments this week.")

    # Returns
    returns = get_recent_returns(data, 7)
    lines.append(f"\n## 3. Laptop Returns This Week ({len(returns)})\n")
    if returns:
        lines.append("| Employee | Laptop | Reason |")
        lines.append("|----------|--------|--------|")
        for r in returns:
            lines.append(f"| {r.get('Username', 'N/A')} | {r.get('Laptop Make', '')} {r.get('Laptop Model', '')} | {r.get('Resigned/Replacement', '')} |")

    # Aging — full table
    aging = get_aging_laptops(data)
    lines.append(f"\n## 4. Aging Analysis ({len(aging)} laptops > 3.5 years)\n")
    if aging:
        lines.append("| Employee | Asset Tag | Make/Model | Purchase Date | Age (yrs) | Priority |")
        lines.append("|----------|-----------|------------|---------------|-----------|----------|")
        for a in aging:
            lines.append(f"| {a['employee']} | {a['tag']} | {a['make']} {a['model']} | {a['start_date']} | {a['age_years']} | {a['priority']} |")

    # Age distribution
    dist = get_age_distribution(data)
    lines.append("\n### Age Distribution\n")
    lines.append("| Bracket | Count |")
    lines.append("|---------|-------|")
    for bracket, count in dist.items():
        lines.append(f"| {bracket} | {count} |")

    # Laptop Procurement
    laptop_spend = get_laptop_spend(data)
    purchases = get_recent_purchases(data, 30)
    lines.append(f"\n## 5. Laptop Procurement — {TODAY.strftime('%B %Y')}\n")
    if laptop_spend["models"]:
        lines.append(f"**Joiners this month:** {laptop_spend['total_joiners']}  ")
        lines.append(f"**Laptop spend this month:** {fmt_usd(laptop_spend['total_spend'])}\n")
        lines.append("| Model | Joiners | Spend |")
        lines.append("|-------|---------|-------|")
        for m in laptop_spend["models"]:
            lines.append(f"| {m['model']} | {m['joiners']} | {fmt_usd(m['spend'])} |")
    else:
        lines.append("No laptop procurement data for this month.\n")
    if purchases:
        lines.append(f"\n### New Laptops Purchased ({len(purchases)})\n")
        lines.append("| Brand | Model | Serial | Purchase Date |")
        lines.append("|-------|-------|--------|---------------|")
        for p in purchases:
            lines.append(f"| {p['brand']} | {p['model']} | {p['serial']} | {p['date'].strftime('%d %b %Y')} |")

    # App Spend
    total_spend, renewals = get_current_month_spend(data)
    lines.append(f"\n## 6. App Spend — {TODAY.strftime('%B %Y')}\n")
    lines.append(f"**Total this month:** {fmt_usd(total_spend)}\n")
    if renewals:
        lines.append("### Upcoming Renewals (next 30 days)\n")
        lines.append("| Application | Department | Renewal Date | Frequency |")
        lines.append("|-------------|------------|--------------|-----------|")
        for r in renewals:
            lines.append(f"| {r['app']} | {r['dept']} | {r['date'].strftime('%d %b %Y')} | {r['frequency']} |")

    # Joiners
    joiners = get_upcoming_joiners(data, 30)
    lines.append(f"\n## 7. Upcoming Joiners (next 30 days) — {len(joiners)}\n")
    if joiners:
        lines.append("| Name | Department | Designation | DOJ |")
        lines.append("|------|------------|-------------|-----|")
        for j in joiners:
            lines.append(f"| {j['name']} | {j['department']} | {j['designation']} | {j['doj'].strftime('%d %b %Y')} |")

    # Summary stats
    total_assigned = len(data["assigned"])
    total_stock = len(data["in_stock"])
    avg_age = 0
    age_count = 0
    for row in data["assigned"]:
        dt = parse_date(row.get("Warranty Start Date"))
        if dt:
            avg_age += age_years(dt)
            age_count += 1
    avg_age = round(avg_age / age_count, 1) if age_count else 0

    lines.append("\n## 8. Summary\n")
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| Total Laptops Assigned | {total_assigned} |")
    lines.append(f"| Laptops Available | {total_stock} |")
    lines.append(f"| Backup Laptops (3yr+) | {len(data['backup'])} |")
    lines.append(f"| Average Laptop Age | {avg_age} years |")
    lines.append(f"| Laptops > 3.5yr | {len(aging)} |")
    lines.append(f"| Laptop Spend This Month | {fmt_usd(laptop_spend['total_spend'])} |")
    lines.append(f"| App Spend This Month | {fmt_usd(total_spend)} |")
    lines.append(f"| Upcoming Joiners (30d) | {len(joiners)} |")

    lines.append(f"\n---\n_Generated: {TODAY.strftime('%d %B %Y')}_")
    return "\n".join(lines)


def generate_monthly_slack(data: dict) -> str:
    lines = [f"*📊 IT Monthly Report — {TODAY.strftime('%B %Y')}*\n"]

    total_assigned = len(data["assigned"])
    aging = get_aging_laptops(data)
    assignments_month = get_recent_assignments(data, 30)
    replacements = [a for a in assignments_month if str(a.get("New Joiner/Replacement", "")).lower() == "replacement"]

    # 1. Highlights
    lines.append("*1. Monthly Highlights*")
    lines.append(f"• Total laptops assigned: {total_assigned}")
    lines.append(f"• New assignments this month: {len(assignments_month)}")
    lines.append(f"• Replacements done: {len(replacements)}")
    lines.append(f"• Laptops flagged for replacement (>3.5yr): {len(aging)}")

    # 2. Stock Health
    stock = get_stock_summary(data)
    lines.append("\n*2. Stock Health*")
    for item, count in stock.items():
        icon = "🟢" if count > 5 else ("🟡" if count >= 2 else "🔴")
        lines.append(f"• {icon} {item}: {count}")

    # 3. Aging
    dist = get_age_distribution(data)
    lines.append("\n*3. Aging Overview*")
    for bracket, count in dist.items():
        lines.append(f"• {bracket}: {count}")

    # 4. Laptop Procurement
    laptop_spend = get_laptop_spend(data)
    purchases = get_recent_purchases(data, 30)
    lines.append(f"\n*4. Laptop Procurement — {TODAY.strftime('%B %Y')}*")
    if laptop_spend["models"]:
        lines.append(f"• Joiners this month: {laptop_spend['total_joiners']}")
        lines.append(f"• Laptop spend this month: {fmt_usd(laptop_spend['total_spend'])}")
        for m in laptop_spend["models"]:
            lines.append(f"  - {m['model']}: {m['joiners']} joiners, {fmt_usd(m['spend'])}")
    else:
        lines.append("• No laptop procurement data for this month")
    if purchases:
        lines.append(f"• New laptops procured this month: {len(purchases)}")

    # 5. App Spend
    total_spend, renewals = get_current_month_spend(data)
    lines.append(f"\n*5. App Spend — {TODAY.strftime('%B %Y')}*")
    lines.append(f"• Total app spend this month: {fmt_usd(total_spend)}")
    lines.append(f"• Renewals in next 30 days: {len(renewals)}")

    # 6. Procurement Recommendation
    stock_laptops = stock["Laptops (ready)"]
    joiners_30 = get_upcoming_joiners(data, 30)
    joiners_90 = get_upcoming_joiners(data, 90)
    aging_critical = [a for a in aging if a["priority"] == "Critical"]
    need = len(joiners_30) + len(aging_critical) - stock_laptops
    lines.append("\n*6. Procurement Recommendation*")
    if need > 0:
        lines.append(f"• ⚠️ Order {need} laptops: {len(joiners_30)} joiners expected + {len(aging_critical)} critical replacements, only {stock_laptops} in stock")
    else:
        lines.append(f"• ✅ Stock sufficient: {stock_laptops} available for {len(joiners_30)} joiners + {len(aging_critical)} replacements")
    lines.append(f"• 90-day joiner forecast: {len(joiners_90)}")

    # 7. Renewals
    lines.append(f"\n*7. Upcoming Renewals*")
    for r in renewals[:5]:
        lines.append(f"• {r['app']} — {r['date'].strftime('%d %b')}")

    # 8. Joiners
    lines.append(f"\n*8. Joiners & Onboarding* ({len(joiners_30)} this month)")
    for j in joiners_30[:5]:
        lines.append(f"• {j['name']} — {j['department']} ({j['doj'].strftime('%d %b')})")

    lines.append(f"\n_Generated: {TODAY.strftime('%d %B %Y')}_")
    return "\n".join(lines)


def generate_monthly_full(data: dict) -> str:
    """Monthly full report includes everything from weekly + extra sections."""
    # Start with weekly full report content
    lines = [f"# IT Monthly Report — {TODAY.strftime('%B %Y')}\n"]

    # Include all weekly sections
    weekly = generate_weekly_full(data)
    # Skip the weekly header, use the rest
    weekly_lines = weekly.split("\n")[1:]
    lines.extend(weekly_lines)

    # Section: Procurement Plan vs Actual
    lines.append(f"\n## 8. Budget vs Actual\n")
    if data["proc_plan"]:
        lines.append("### Planned Procurement\n")
        lines.append("| Department | Model | Qty | Avg Price | Total |")
        lines.append("|------------|-------|-----|-----------|-------|")
        for row in data["proc_plan"]:
            dept = row.get("Department", row.get("Laptop Procurement Plan & Cost Estimation", ""))
            model = row.get("Model", "")
            qty = row.get("Quantity", "")
            price = row.get("Avg Price/Laptop (INR) (As per current market price)", "")
            total = row.get("Total Price (INR)", "")
            if dept and model:
                lines.append(f"| {dept} | {model} | {qty} | {fmt_usd(price)} | {fmt_usd(total)} |")

    # Onboarding checklist status
    lines.append(f"\n## 9. Onboarding Checklist Status\n")
    if data["checklist"]:
        checklist_cols = ["Email ID Creation", "Reporting Manager Update", "Enable MFA",
                          "Invite on Clickup", "Invite on slack", "Asset policy Acknowledgement"]
        lines.append("| Employee | " + " | ".join(checklist_cols) + " |")
        lines.append("|----------|" + "|".join(["---"] * len(checklist_cols)) + "|")
        for row in data["checklist"][-10:]:  # last 10 joiners
            name = row.get("Name ", row.get("Name", ""))
            vals = []
            for col in checklist_cols:
                v = row.get(col, "")
                vals.append("✅" if v and str(v).strip().lower() not in ("", "none", "no") else "❌")
            if name:
                lines.append(f"| {name} | " + " | ".join(vals) + " |")

    lines.append(f"\n---\n_Generated: {TODAY.strftime('%d %B %Y')}_")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    if len(sys.argv) < 2 or sys.argv[1] not in ("weekly", "monthly"):
        print("Usage: generate-report.py weekly|monthly", file=sys.stderr)
        sys.exit(1)

    report_type = sys.argv[1]
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Loading Excel data from {DATA_DIR} …")
    data = load_data()
    print(f"  Loaded: {len(data['assigned'])} assigned laptops, {len(data['history'])} history records, "
          f"{len(data['spend'])} spend rows, {len(data['joinings'])} joiners")

    if report_type == "weekly":
        slack = generate_weekly_slack(data)
        full = generate_weekly_full(data)
    else:
        slack = generate_monthly_slack(data)
        full = generate_monthly_full(data)

    (OUTPUT_DIR / "slack-summary.md").write_text(slack, encoding="utf-8")
    (OUTPUT_DIR / "full-report.md").write_text(full, encoding="utf-8")

    print(f"Reports saved to {OUTPUT_DIR}/")
    print(f"  slack-summary.md: {len(slack)} chars")
    print(f"  full-report.md: {len(full)} chars")


if __name__ == "__main__":
    main()
