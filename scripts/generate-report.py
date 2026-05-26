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

import json
import sys
import pathlib
import datetime as dt
from datetime import timedelta
from collections import defaultdict
from typing import Optional

import openpyxl

DATA_DIR = pathlib.Path(__file__).resolve().parent.parent / "data"
OUTPUT_DIR = pathlib.Path(__file__).resolve().parent.parent / "output"
SNAPSHOT_DIR = pathlib.Path(__file__).resolve().parent.parent / "snapshots"
TODAY = dt.datetime.now().date()
AGE_THRESHOLD_DAYS = int(3.5 * 365)  # 1277 days


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def is_truly_assigned(row: dict) -> bool:
    """True if the row represents an actively assigned laptop.
    Excludes rows where Employee ID is blank or marked "In Stock" /
    "Hand over" / similar non-assignment placeholders."""
    emp_id = row.get("Employee ID")
    if emp_id is None:
        return False
    emp_id_str = str(emp_id).strip().lower()
    if not emp_id_str:
        return False
    # Common non-assignment markers in the Employee ID column
    non_assignment = ("in stock", "instock", "hand over", "handover", "stock")
    if any(marker in emp_id_str for marker in non_assignment):
        return False
    return True


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


def fmt_inr(amount) -> str:
    try:
        n = float(amount)
    except (TypeError, ValueError):
        return "N/A"
    if n >= 10_000_000:
        return f"₹{n/10_000_000:,.2f} Cr"
    if n >= 100_000:
        return f"₹{n/100_000:,.2f} L"
    return f"₹{n:,.0f}"


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
        "configuration": read_sheet(proc_wb, "Configuration"),
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
        if dt and cutoff <= dt <= TODAY:
            results.append(row)
    return results


def get_assignments_in_window(data: dict, start_days_ago: int, end_days_ago: int) -> list[dict]:
    """Assignments with date in [TODAY - start_days_ago, TODAY - end_days_ago]."""
    start = TODAY - timedelta(days=start_days_ago)
    end = TODAY - timedelta(days=end_days_ago)
    results = []
    for row in data["history"]:
        d = parse_date(row.get("Assigned Date"))
        if d and start <= d <= end:
            results.append(row)
    return results


def get_returns_in_window(data: dict, start_days_ago: int, end_days_ago: int) -> list[dict]:
    start = TODAY - timedelta(days=start_days_ago)
    end = TODAY - timedelta(days=end_days_ago)
    results = []
    for row in data["returned"]:
        d = parse_date(row.get("Returned Date"))
        if d and start <= d <= end:
            results.append(row)
    return results


def get_recent_returns(data: dict, days: int = 7) -> list[dict]:
    cutoff = TODAY - timedelta(days=days)
    results = []
    for row in data["returned"]:
        dt = parse_date(row.get("Returned Date"))
        if dt and cutoff <= dt <= TODAY:
            results.append(row)
    return results


def get_weekly_activity_comparison(data: dict) -> dict:
    """Compare this-week (last 7d) vs previous-week (7-14d ago) activity."""
    this_assigns = get_assignments_in_window(data, 7, 0)
    prev_assigns = get_assignments_in_window(data, 14, 8)
    this_repl = [a for a in this_assigns if str(a.get("New Joiner/Replacement", "")).lower() == "replacement"]
    prev_repl = [a for a in prev_assigns if str(a.get("New Joiner/Replacement", "")).lower() == "replacement"]
    this_returns = get_returns_in_window(data, 7, 0)
    prev_returns = get_returns_in_window(data, 14, 8)
    return {
        "assignments": {"this": len(this_assigns), "prev": len(prev_assigns)},
        "new_joiner_assigns": {
            "this": sum(1 for a in this_assigns if str(a.get("New Joiner/Replacement", "")).lower() == "new joiner"),
            "prev": sum(1 for a in prev_assigns if str(a.get("New Joiner/Replacement", "")).lower() == "new joiner"),
        },
        "replacements": {"this": len(this_repl), "prev": len(prev_repl)},
        "returns": {"this": len(this_returns), "prev": len(prev_returns)},
    }


# ---------------------------------------------------------------------------
# Snapshot helpers (for stock/week-over-week comparison)
# ---------------------------------------------------------------------------

def get_procurement_runway(data: dict, lookback_weeks: int = 4) -> dict:
    """Estimate weeks of laptop stock at current assignment rate.

    Uses new-joiner assignments over the past `lookback_weeks` weeks as the
    consumption signal (replacements pull from stock too, but new joiners are
    the dominant driver). Returns dict with avg_per_week, stock_ready, weeks.
    """
    cutoff = TODAY - timedelta(days=lookback_weeks * 7)
    count = 0
    for row in data["history"]:
        d = parse_date(row.get("Assigned Date"))
        if d and cutoff <= d <= TODAY:
            atype = str(row.get("New Joiner/Replacement", "")).lower()
            if "joiner" in atype:
                count += 1
    avg_per_week = count / lookback_weeks if lookback_weeks else 0
    stock_ready = len(data["in_stock"])
    weeks: Optional[float] = None
    if avg_per_week > 0:
        weeks = round(stock_ready / avg_per_week, 1)
    return {
        "avg_per_week": round(avg_per_week, 1),
        "stock_ready": stock_ready,
        "weeks": weeks,
    }


def get_spend_pace(data: dict) -> dict:
    """Compare current-month laptop spend (actual) to planned budget.

    Pulls planned total from procurement_plan "Laptop procurement plan" sheet.
    Returns dict with planned, actual, pct_used (or None if planned=0).
    """
    laptop_spend = get_laptop_spend(data)
    actual = laptop_spend["total_spend"]
    planned = 0.0
    for row in data.get("proc_plan", []):
        v = row.get("Total Price (INR)")
        if isinstance(v, (int, float)):
            planned += float(v)
    # Planned is annual; divide by 12 for per-month budget
    monthly_planned = planned / 12 if planned else 0
    pct = None
    if monthly_planned > 0:
        pct = round((actual / monthly_planned) * 100, 0)
    return {
        "actual": actual,
        "monthly_planned": monthly_planned,
        "annual_planned": planned,
        "pct_used": pct,
    }


def get_health_check(data: dict, prev_snap: Optional[dict] = None) -> dict:
    """Compute traffic-light status for top dashboard tiles."""
    aging = get_aging_laptops(data)
    critical = sum(1 for a in aging if a["priority"] == "Critical")
    runway = get_procurement_runway(data)
    joiners_7 = get_upcoming_joiners(data, 7)
    stock_ready = len(data["in_stock"])
    pace = get_spend_pace(data)

    # Stock: based on runway weeks
    if runway["weeks"] is None:
        stock_status = "🟢"  # no recent consumption
    elif runway["weeks"] >= 4:
        stock_status = "🟢"
    elif runway["weeks"] >= 2:
        stock_status = "🟡"
    else:
        stock_status = "🔴"

    # Aging: based on critical count
    if critical == 0:
        aging_status = "🟢"
    elif critical <= 5:
        aging_status = "🟡"
    else:
        aging_status = "🔴"

    # Joiner prep: stock vs joiners next 7 days
    if not joiners_7:
        joiner_status = "🟢"
    elif stock_ready >= len(joiners_7):
        joiner_status = "🟢"
    elif stock_ready >= len(joiners_7) - 2:
        joiner_status = "🟡"
    else:
        joiner_status = "🔴"

    # Spend: vs monthly budget pace
    if pace["pct_used"] is None:
        spend_status = "🟢"
    elif pace["pct_used"] < 90:
        spend_status = "🟢"
    elif pace["pct_used"] < 110:
        spend_status = "🟡"
    else:
        spend_status = "🔴"

    return {
        "stock": stock_status,
        "aging": aging_status,
        "joiner_prep": joiner_status,
        "spend": spend_status,
        "_critical_aging": critical,
        "_runway_weeks": runway["weeks"],
        "_joiners_7": len(joiners_7),
        "_stock_ready": stock_ready,
        "_pace_pct": pace["pct_used"],
    }


def get_risk_callouts(data: dict, hc: dict) -> list[str]:
    """Auto-detect risks worth highlighting at the top of the report."""
    risks: list[str] = []
    runway = hc["_runway_weeks"]
    joiners_7 = hc["_joiners_7"]
    stock_ready = hc["_stock_ready"]
    critical = hc["_critical_aging"]
    pace_pct = hc["_pace_pct"]

    # Joiner stock shortage
    if joiners_7 > stock_ready:
        gap = joiners_7 - stock_ready
        risks.append(f"🔴 *{joiners_7} joiners next week*, only {stock_ready} laptops in stock — order *{gap}* laptops")
    elif joiners_7 and stock_ready - joiners_7 <= 2:
        risks.append(f"🟡 Tight stock: {stock_ready} laptops vs {joiners_7} joiners next week")

    # Stock runway
    if runway is not None and runway < 2:
        risks.append(f"🔴 *Stock runway: {runway} weeks* at current assignment pace")
    elif runway is not None and runway < 4:
        risks.append(f"🟡 Stock runway: {runway} weeks — plan procurement soon")

    # Critical aging
    if critical > 5:
        risks.append(f"🟡 *{critical} laptops >4 yrs* still assigned — schedule replacements")

    # Spend pace
    if pace_pct is not None and pace_pct > 110:
        risks.append(f"🔴 Laptop spend pace: *{pace_pct:.0f}%* of monthly budget — over by {pace_pct - 100:.0f}%")

    return risks


def current_snapshot(data: dict) -> dict:
    """Build a snapshot of key metrics for persistence."""
    total_assigned = sum(1 for r in data["assigned"] if is_truly_assigned(r))
    aging = get_aging_laptops(data)
    laptop_spend = get_laptop_spend(data)
    total_app, _, _, _ = get_current_month_spend(data)
    joiners_30 = get_upcoming_joiners(data, 30)
    return {
        "date": TODAY.isoformat(),
        "stock_ready": len(data["in_stock"]),
        "stock_backup": len(data["backup"]),
        "total_assigned": total_assigned,
        "total_laptops": total_assigned + len(data["in_stock"]) + len(data["backup"]),
        "aging_count": len(aging),
        "laptop_spend_month": laptop_spend["total_spend"],
        "laptop_joiners_month": laptop_spend["total_joiners"],
        "app_spend_month": total_app,
        "joiners_next_30": len(joiners_30),
    }


def load_previous_snapshot() -> Optional[dict]:
    """Read the snapshot saved by the previous run, if any."""
    path = SNAPSHOT_DIR / "latest.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def save_snapshot(snap: dict) -> None:
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    (SNAPSHOT_DIR / "latest.json").write_text(
        json.dumps(snap, indent=2), encoding="utf-8"
    )


def _delta_icon(cur: float, prev: float, good_is_up: bool = False) -> str:
    if cur == prev:
        return "➖"
    up = cur > prev
    if good_is_up:
        return "🟢 ↑" if up else "🔴 ↓"
    return "🔴 ↑" if up else "🟢 ↓"


def _fmt_delta(cur: float, prev: Optional[float], as_int: bool = True) -> str:
    if prev is None:
        return "—"
    diff = cur - prev
    if diff == 0:
        return "±0"
    sign = "+" if diff > 0 else ""
    if as_int:
        return f"{sign}{int(diff)}"
    return f"{sign}{diff:,.0f}"


def get_aging_laptops(data: dict) -> list[dict]:
    """Return assigned laptops older than 3.5 years, sorted oldest first."""
    aging = []
    for row in data["assigned"]:
        if not is_truly_assigned(row):
            continue
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
        if not is_truly_assigned(row):
            continue
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
    total_row = None
    for row in data["actual_spend"]:
        model = row.get("Model") or row.get("col_0", "")
        model_str = str(model).strip().lower() if model else ""
        # Capture the Total row separately for authoritative monthly spend
        if model_str in ("total", "grand total"):
            total_row = row
            continue
        if not model_str or model_str in ("none",):
            continue

        # Per-model: pull joiners for current month (individual spend values
        # in this sheet are unreliable; use Total row below for actual INR)
        joiners = 0
        for key, val in row.items():
            key_lower = str(key).strip().lower()
            for abbr in abbrevs:
                if abbr in key_lower and "joiner" in key_lower:
                    try:
                        joiners = int(val) if val not in (None, "") else 0
                    except (TypeError, ValueError):
                        joiners = 0

        if joiners:
            result["models"].append({
                "model": str(model).strip(),
                "joiners": joiners,
            })
            result["total_joiners"] += joiners

    # Authoritative spend from the Total row
    if total_row:
        for key, val in total_row.items():
            key_lower = str(key).strip().lower()
            for abbr in abbrevs:
                if abbr in key_lower and "spend" in key_lower:
                    try:
                        result["total_spend"] = float(val) if val not in (None, "") else 0.0
                    except (TypeError, ValueError):
                        result["total_spend"] = 0.0

    return result


def get_recent_purchases(data: dict, days: int = 30) -> list[dict]:
    """Get laptops purchased within the given number of past days (up to today)."""
    cutoff = TODAY - timedelta(days=days)
    purchases = []
    for row in data["purchased"]:
        d = parse_date(row.get("Warranty Start Date"))
        if d and cutoff <= d <= TODAY:
            purchases.append({
                "brand": row.get("Brand", ""),
                "model": row.get("Model", ""),
                "serial": row.get("Serial no", ""),
                "date": d,
            })
    purchases.sort(key=lambda x: x["date"], reverse=True)
    return purchases


def get_purchases_this_month(data: dict) -> list[dict]:
    """Get laptops purchased within the current calendar month."""
    purchases = []
    for row in data["purchased"]:
        d = parse_date(row.get("Warranty Start Date"))
        if d and d.year == TODAY.year and d.month == TODAY.month and d <= TODAY:
            purchases.append({
                "brand": row.get("Brand", ""),
                "model": row.get("Model", ""),
                "serial": row.get("Serial no", ""),
                "date": d,
            })
    purchases.sort(key=lambda x: x["date"], reverse=True)
    return purchases


# Row names in spend tracker that are laptop/hardware costs, not app subscriptions
# Match the specific hardware-procurement row name(s) — not loose substrings
# like "laptop", which could catch Microsoft Surface Laptop etc.
HARDWARE_SPEND_KEYWORDS = ["laptops procurement"]
# Row names that are aggregate/total rows (would double-count if summed)
TOTAL_ROW_KEYWORDS = ["total", "grand total", "sum"]


def _is_hardware_row(row: dict) -> bool:
    app_name = str(row.get("APPLICATION / SW / LICENSE", "")).lower()
    return any(kw in app_name for kw in HARDWARE_SPEND_KEYWORDS)


def _is_total_row(row: dict) -> bool:
    app_name = str(row.get("APPLICATION / SW / LICENSE", "")).strip().lower()
    return app_name in TOTAL_ROW_KEYWORDS


def get_current_month_spend(data: dict) -> tuple[float, list[dict], float, float]:
    """Get app spend for current month, upcoming renewals, hardware spend, and grand total.

    Returns: (app_only_total, renewals, hardware_total, grand_total)
        - app_only_total: sum of app/subscription rows (excludes hardware + Total rows)
        - hardware_total: sum of "Laptops Procurement", "Antivirus,MDM" type rows
        - grand_total: app + hardware (matches the Total row in the sheet)
    """
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

    app_total = 0.0
    hw_total = 0.0
    if month_key:
        for row in data["spend"]:
            if _is_total_row(row):
                continue
            val = row.get(month_key)
            if not (val and isinstance(val, (int, float))):
                continue
            if _is_hardware_row(row):
                hw_total += float(val)
            else:
                app_total += float(val)

    # Upcoming renewals (app only)
    renewals = []
    cutoff = TODAY + timedelta(days=30)
    for row in data["spend"]:
        if _is_hardware_row(row) or _is_total_row(row):
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
    return app_total, renewals, hw_total, app_total + hw_total


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


def lookup_laptop_config(data: dict, department: str, designation: str) -> str:
    """Find the standard laptop config for a joiner based on department/designation.

    Returns a short string like "Lenovo L14 / 16GB / i5" or "" if no match.
    """
    if not department and not designation:
        return ""
    dept_l = str(department).strip().lower()
    desig_l = str(designation).strip().lower()
    best_match = None
    for row in data.get("configuration", []):
        cfg_dept = str(row.get("Department & Owner", "")).strip().lower()
        cfg_role = str(row.get("Role / Position", "")).strip().lower()
        # Match if department contains cfg_dept or vice versa, or role matches
        dept_hit = cfg_dept and (cfg_dept in dept_l or dept_l in cfg_dept)
        role_hit = cfg_role and (cfg_role in desig_l or desig_l in cfg_role)
        if dept_hit and role_hit:
            best_match = row
            break  # exact match on both
        if dept_hit and not best_match:
            best_match = row
    if not best_match:
        return ""
    device = str(best_match.get("Device Type", "")).strip()
    ram = str(best_match.get("RAM", "")).strip()
    proc = str(best_match.get("Processor", "")).strip()
    parts = [p for p in (device, ram, proc) if p and p.lower() != "none"]
    return " / ".join(parts)


def get_joiners_with_laptop_needs(data: dict, days: int = 7) -> list[dict]:
    """Upcoming joiners in the next N days with their required laptop config."""
    joiners = get_upcoming_joiners(data, days)
    for j in joiners:
        j["laptop_config"] = lookup_laptop_config(data, j["department"], j["designation"])
        j["days_until"] = (j["doj"] - TODAY).days
    return joiners


def get_stock_vs_joiners(data: dict, days: int = 7) -> dict:
    """Compare current laptop stock to upcoming joiner demand.

    Returns a dict with keys:
      stock_ready, stock_backup, joiners_next_week, joiners_next_30_days,
      gap_next_week, gap_next_30_days, status ("ok"/"watch"/"short")
    """
    stock_ready = len(data["in_stock"])
    stock_backup = len(data["backup"])
    joiners_7 = get_upcoming_joiners(data, days)
    joiners_30 = get_upcoming_joiners(data, 30)
    gap_7 = len(joiners_7) - stock_ready
    gap_30 = len(joiners_30) - stock_ready
    if gap_7 > 0:
        status = "short"
    elif gap_30 > 0:
        status = "watch"
    else:
        status = "ok"
    return {
        "stock_ready": stock_ready,
        "stock_backup": stock_backup,
        "joiners_next_week": len(joiners_7),
        "joiners_next_30_days": len(joiners_30),
        "gap_next_week": gap_7,
        "gap_next_30_days": gap_30,
        "status": status,
    }


# ---------------------------------------------------------------------------
# Report generators
# ---------------------------------------------------------------------------

def generate_weekly_slack(data: dict, prev_snap: Optional[dict] = None) -> str:
    """Concise at-a-glance dashboard. Full breakdown lives in the Word doc."""
    lines = [f"*📊 IT Weekly Report — {TODAY.strftime('%d %b %Y')}*"]

    hc = get_health_check(data, prev_snap)
    risks = get_risk_callouts(data, hc)
    aging = get_aging_laptops(data)
    critical_aging = [a for a in aging if a["priority"] == "Critical"]
    laptop_spend = get_laptop_spend(data)
    app_total, renewals, hw_spend, _ = get_current_month_spend(data)
    runway = get_procurement_runway(data)
    joiners_7 = get_upcoming_joiners(data, 7)

    # ── 1. HEALTH CHECK ──
    lines.append(f"\n*🚦 Health Check*")
    lines.append(
        f"Stock {hc['stock']} · Aging {hc['aging']} · "
        f"Joiner Prep {hc['joiner_prep']} · Spend {hc['spend']}"
    )

    # ── 2. RISKS & ACTIONS ──
    if risks:
        lines.append(f"\n*⚠️ Risks & Actions*")
        for r in risks:
            lines.append(f"• {r}")

    # ── 3. KEY NUMBERS (vs last week) ──
    lines.append(f"\n*📊 Key Numbers*")
    stock_ready = hc["_stock_ready"]
    prev_ready = (prev_snap or {}).get("stock_ready")
    prev_aging = (prev_snap or {}).get("aging_count")
    prev_app = (prev_snap or {}).get("app_spend_month")
    prev_lap = (prev_snap or {}).get("laptop_spend_month")

    def _kn(label: str, cur, prev, fmt=lambda x: str(x)) -> str:
        if prev is None:
            return f"• {label}: *{fmt(cur)}*"
        d = _fmt_delta(cur, prev, as_int=not isinstance(cur, float) or cur.is_integer())
        return f"• {label}: *{fmt(cur)}* ({d})"

    lines.append(_kn("Stock ready", stock_ready, prev_ready))
    lines.append(f"• Joiners next 7d: *{hc['_joiners_7']}* | next 30d: *{len(get_upcoming_joiners(data, 30))}*")
    lines.append(_kn("Aging >3.5yr", len(aging), prev_aging))
    lines.append(_kn("Laptop spend MTD", laptop_spend["total_spend"], prev_lap, fmt=fmt_inr))
    lines.append(_kn("App spend MTD", app_total, prev_app, fmt=fmt_usd))
    if runway["weeks"] is not None:
        lines.append(f"• Procurement runway: *{runway['weeks']} weeks* (avg {runway['avg_per_week']} joiners/wk)")

    # ── 4. JOINERS NEXT WEEK (only if any) ──
    if joiners_7:
        lines.append(f"\n*⏰ Joiners Next Week ({len(joiners_7)})* — prepare laptops")
        joiners_week = get_joiners_with_laptop_needs(data, 7)
        for j in joiners_week[:5]:
            cfg = f" · _{j['laptop_config']}_" if j['laptop_config'] else ""
            days = f"in {j['days_until']}d" if j['days_until'] > 0 else "today"
            lines.append(f"• {j['name']} — {j['department']} (DOJ {j['doj'].strftime('%d %b')}, {days}){cfg}")
        if len(joiners_week) > 5:
            lines.append(f"  _…and {len(joiners_week)-5} more_")

    # ── 5. TOP CRITICAL AGING (top 3 only) ──
    if critical_aging:
        lines.append(f"\n*🔝 Top Critical Aging* (showing 3 of {len(critical_aging)})")
        for a in critical_aging[:3]:
            lines.append(f"• {a['employee']} — {a['make']} {a['model']} ({a['age_years']}yr)")

    # ── 6. RENEWALS (only if any) ──
    if renewals:
        lines.append(f"\n*🔔 Renewals next 30d* ({len(renewals)})")
        for r in renewals[:3]:
            lines.append(f"• {r['app']} ({r['date'].strftime('%d %b')})")

    lines.append(f"\n_Full breakdown in the attached Word doc. React 👍/👎 or reply with feedback._")
    return "\n".join(lines)


def generate_joiner_alert(data: dict) -> str:
    """Separate Slack alert: joiners next week + stock vs joiners analysis."""
    lines = [f"*🚨 Joiner Alert & Stock Check — {TODAY.strftime('%d %B %Y')}*\n"]

    # Joiners next week with laptop config
    joiners_week = get_joiners_with_laptop_needs(data, 7)
    lines.append(f"*⏰ Joiners Next Week* ({len(joiners_week)}) — prepare laptops")
    if joiners_week:
        for j in joiners_week:
            cfg = f" · _{j['laptop_config']}_" if j['laptop_config'] else ""
            days = f"in {j['days_until']}d" if j['days_until'] > 0 else "today"
            lines.append(f"• {j['name']} — {j['department']}, {j['designation']} (DOJ {j['doj'].strftime('%d %b')}, {days}){cfg}")
    else:
        lines.append("• No joiners in the next 7 days ✅")

    # Stock vs Joiners
    svj = get_stock_vs_joiners(data, 7)
    icon = {"ok": "✅", "watch": "🟡", "short": "🔴"}[svj["status"]]
    lines.append(f"\n*{icon} Stock vs Joiners*")
    lines.append(f"• Laptops in stock: {svj['stock_ready']} (ready) + {svj['stock_backup']} (backup)")
    lines.append(f"• Joiners next 7 days: {svj['joiners_next_week']} | next 30 days: {svj['joiners_next_30_days']}")
    if svj["gap_next_week"] > 0:
        lines.append(f"• 🔴 *Short {svj['gap_next_week']} laptop(s)* for next week's joiners — arrange immediately!")
    elif svj["gap_next_30_days"] > 0:
        lines.append(f"• 🟡 Will be short {svj['gap_next_30_days']} laptop(s) within 30 days — plan procurement")
    else:
        lines.append(f"• ✅ Stock covers next 30 days of joiners")

    return "\n".join(lines)


def generate_weekly_full(data: dict, prev_snap: Optional[dict] = None) -> str:
    lines = [f"# IT Weekly Report — {TODAY.strftime('%d %B %Y')}\n"]

    # ── Dashboard preamble ──
    hc = get_health_check(data, prev_snap)
    risks = get_risk_callouts(data, hc)
    runway = get_procurement_runway(data)
    pace = get_spend_pace(data)

    lines.append("## 🚦 Health Check\n")
    lines.append("| Area | Status |")
    lines.append("|------|--------|")
    lines.append(f"| Stock | {hc['stock']} |")
    lines.append(f"| Aging | {hc['aging']} |")
    lines.append(f"| Joiner Prep | {hc['joiner_prep']} |")
    lines.append(f"| Spend | {hc['spend']} |")

    if risks:
        lines.append("\n## ⚠️ Risks & Actions\n")
        for r in risks:
            # Strip Slack-style bold (* → markdown __ would italic; keep plain)
            lines.append(f"- {r.replace('*', '')}")

    lines.append("\n## 📊 Headline Metrics\n")
    lines.append("| Metric | Value |")
    lines.append("|--------|-------|")
    lines.append(f"| Procurement runway | {runway['weeks']} weeks (avg {runway['avg_per_week']} joiners/wk over 4w) |")
    lines.append(f"| Critical aging laptops | {hc['_critical_aging']} |")
    lines.append(f"| Joiners next 7 days | {hc['_joiners_7']} |")
    lines.append(f"| Stock ready | {hc['_stock_ready']} |")
    if pace["pct_used"] is not None:
        lines.append(f"| Laptop spend pace | {pace['pct_used']:.0f}% of monthly budget ({fmt_inr(pace['actual'])} / {fmt_inr(pace['monthly_planned'])}) |")

    # Stock (side-by-side with previous snapshot)
    stock = get_stock_summary(data)
    lines.append("\n## 1. Stock Levels\n")
    if prev_snap and prev_snap.get("date"):
        prev_date = prev_snap["date"]
        cur_ready = stock.get("Laptops (ready)", 0)
        cur_backup = stock.get("Laptops (3yr+ backup)", 0)
        prev_ready = prev_snap.get("stock_ready", cur_ready)
        prev_backup = prev_snap.get("stock_backup", cur_backup)
        lines.append(f"| Asset Type | This Week | Last Week ({prev_date}) | Δ |")
        lines.append("|------------|-----------|------------------------|---|")
        lines.append(f"| Laptops (ready) | {cur_ready} | {prev_ready} | {_fmt_delta(cur_ready, prev_ready)} |")
        lines.append(f"| Laptops (3yr+ backup) | {cur_backup} | {prev_backup} | {_fmt_delta(cur_backup, prev_backup)} |")
    else:
        lines.append("| Asset Type | Available |")
        lines.append("|------------|-----------|")
        for item, count in stock.items():
            lines.append(f"| {item} | {count} |")

    # Activity: This Week vs Last Week
    cmp = get_weekly_activity_comparison(data)
    lines.append("\n## 1a. Activity: This Week vs Last Week\n")
    lines.append("| Metric | This Week | Last Week | Δ |")
    lines.append("|--------|-----------|-----------|---|")
    for label, key in [
        ("New assignments", "assignments"),
        ("New joiner assignments", "new_joiner_assigns"),
        ("Replacements", "replacements"),
        ("Returns", "returns"),
    ]:
        this_v = cmp[key]["this"]
        prev_v = cmp[key]["prev"]
        lines.append(f"| {label} | {this_v} | {prev_v} | {_fmt_delta(this_v, prev_v)} |")

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
    purchases = get_purchases_this_month(data)
    procured_full = laptop_spend["total_joiners"] if laptop_spend["total_spend"] > 0 else 0
    lines.append(f"\n## 5. Laptop Procurement — {TODAY.strftime('%B %Y')}\n")
    if laptop_spend["total_joiners"] or laptop_spend["total_spend"]:
        lines.append("### Summary\n")
        lines.append("| Metric | Value |")
        lines.append("|--------|-------|")
        lines.append(f"| Joiners this month | {laptop_spend['total_joiners']} |")
        lines.append(f"| Laptop spend this month | {fmt_inr(laptop_spend['total_spend'])} |")
        lines.append(f"| Laptops procured this month | {procured_full} |")
        if procured_full > 0:
            lines.append("")
            lines.append("### Breakdown by Model\n")
            lines.append("| Model | Joiners |")
            lines.append("|-------|---------|")
            for m in laptop_spend["models"]:
                lines.append(f"| {m['model']} | {m['joiners']} |")
    else:
        lines.append("No laptop procurement data for this month.\n")
    if purchases:
        lines.append(f"\n### New Laptops Purchased ({len(purchases)})\n")
        lines.append("| Brand | Model | Serial | Purchase Date |")
        lines.append("|-------|-------|--------|---------------|")
        for p in purchases:
            lines.append(f"| {p['brand']} | {p['model']} | {p['serial']} | {p['date'].strftime('%d %b %Y')} |")

    # App Spend
    total_spend, renewals, hw_spend, grand_total = get_current_month_spend(data)
    lines.append(f"\n## 6. App Spend — {TODAY.strftime('%B %Y')}\n")
    lines.append("| Category | Amount |")
    lines.append("|----------|--------|")
    lines.append(f"| Apps / Subscriptions | {fmt_usd(total_spend)} |")
    lines.append(f"| Hardware (laptops/antivirus, excluded from app spend) | {fmt_usd(hw_spend)} |")
    lines.append(f"| **Sheet total** | **{fmt_usd(grand_total)}** |")
    lines.append("")
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

    # Joiners Next Week — Laptop needs
    joiners_week = get_joiners_with_laptop_needs(data, 7)
    lines.append(f"\n### ⏰ Joiners Next Week ({len(joiners_week)}) — Laptop Prep Reminder\n")
    if joiners_week:
        lines.append("| Name | Department | Designation | DOJ | Days | Laptop Config |")
        lines.append("|------|------------|-------------|-----|------|---------------|")
        for j in joiners_week:
            days = f"{j['days_until']}d" if j['days_until'] > 0 else "today"
            lines.append(f"| {j['name']} | {j['department']} | {j['designation']} | {j['doj'].strftime('%d %b %Y')} | {days} | {j['laptop_config'] or '—'} |")
    else:
        lines.append("No joiners in the next 7 days.")

    # Stock vs Joiners Analysis
    svj = get_stock_vs_joiners(data, 7)
    icon = {"ok": "✅", "watch": "🟡", "short": "🔴"}[svj["status"]]
    lines.append(f"\n### {icon} Stock vs Joiners Analysis\n")
    lines.append("| Metric | Value |")
    lines.append("|--------|-------|")
    lines.append(f"| Laptops ready in stock | {svj['stock_ready']} |")
    lines.append(f"| Backup laptops (3yr+) | {svj['stock_backup']} |")
    lines.append(f"| Joiners next 7 days | {svj['joiners_next_week']} |")
    lines.append(f"| Joiners next 30 days | {svj['joiners_next_30_days']} |")
    lines.append(f"| Gap next 7 days | {svj['gap_next_week']} |")
    lines.append(f"| Gap next 30 days | {svj['gap_next_30_days']} |")
    if svj["gap_next_week"] > 0:
        lines.append(f"\n**🔴 Action needed:** short {svj['gap_next_week']} laptop(s) for next week's joiners — arrange immediately.")
    elif svj["gap_next_30_days"] > 0:
        lines.append(f"\n**🟡 Plan procurement:** will be short {svj['gap_next_30_days']} laptop(s) within 30 days.")
    else:
        lines.append(f"\n**✅ Stock sufficient** for next 30 days of joiners.")

    # Summary stats (filter out In Stock / blank rows)
    assigned_rows = [r for r in data["assigned"] if is_truly_assigned(r)]
    total_assigned = len(assigned_rows)
    total_stock = len(data["in_stock"])
    avg_age = 0
    age_count = 0
    for row in assigned_rows:
        dt = parse_date(row.get("Warranty Start Date"))
        if dt:
            avg_age += age_years(dt)
            age_count += 1
    avg_age = round(avg_age / age_count, 1) if age_count else 0

    total_laptops = total_assigned + total_stock + len(data['backup'])
    lines.append("\n## 8. Summary\n")
    lines.append(f"| Metric | Value |")
    lines.append(f"|--------|-------|")
    lines.append(f"| **Total Laptops** | **{total_laptops}** |")
    lines.append(f"| Total Laptops Assigned | {total_assigned} |")
    lines.append(f"| Laptops Available | {total_stock} |")
    lines.append(f"| Backup Laptops (3yr+) | {len(data['backup'])} |")
    lines.append(f"| Average Laptop Age | {avg_age} years |")
    lines.append(f"| Laptops > 3.5yr | {len(aging)} |")
    lines.append(f"| Laptop Spend This Month | {fmt_inr(laptop_spend['total_spend'])} |")
    lines.append(f"| App Spend This Month | {fmt_usd(total_spend)} |")
    lines.append(f"| Upcoming Joiners (30d) | {len(joiners)} |")

    lines.append(f"\n---\n_Generated: {TODAY.strftime('%d %B %Y')}_")
    return "\n".join(lines)


def generate_monthly_slack(data: dict, prev_snap: Optional[dict] = None) -> str:
    """Concise monthly dashboard."""
    lines = [f"*📊 IT Monthly Report — {TODAY.strftime('%B %Y')}*"]

    hc = get_health_check(data, prev_snap)
    risks = get_risk_callouts(data, hc)
    aging = get_aging_laptops(data)
    critical_aging = [a for a in aging if a["priority"] == "Critical"]
    laptop_spend = get_laptop_spend(data)
    app_total, renewals, _, _ = get_current_month_spend(data)
    pace = get_spend_pace(data)
    total_assigned = sum(1 for r in data["assigned"] if is_truly_assigned(r))
    total_laptops = total_assigned + len(data["in_stock"]) + len(data["backup"])
    assignments_month = get_recent_assignments(data, 30)
    replacements = [a for a in assignments_month if str(a.get("New Joiner/Replacement", "")).lower() == "replacement"]
    joiners_30 = get_upcoming_joiners(data, 30)
    joiners_90 = get_upcoming_joiners(data, 90)

    # ── 1. Health Check ──
    lines.append(f"\n*🚦 Health Check*")
    lines.append(
        f"Stock {hc['stock']} · Aging {hc['aging']} · "
        f"Joiner Prep {hc['joiner_prep']} · Spend {hc['spend']}"
    )

    # ── 2. Risks ──
    if risks:
        lines.append(f"\n*⚠️ Risks & Actions*")
        for r in risks:
            lines.append(f"• {r}")

    # ── 3. Monthly Highlights ──
    lines.append(f"\n*📊 Monthly Highlights*")
    lines.append(f"• Total laptops: *{total_laptops}* (Assigned {total_assigned} · Stock {len(data['in_stock'])} · Backup {len(data['backup'])})")
    lines.append(f"• New assignments this month: *{len(assignments_month)}* ({len(replacements)} replacements)")
    lines.append(f"• Aging >3.5yr: *{len(aging)}* ({hc['_critical_aging']} critical)")
    lines.append(f"• Laptop spend MTD: *{fmt_inr(laptop_spend['total_spend'])}*" +
                 (f" ({pace['pct_used']:.0f}% of budget)" if pace['pct_used'] is not None else ""))
    lines.append(f"• App spend MTD: *{fmt_usd(app_total)}*")

    # ── 4. Procurement Recommendation ──
    aging_critical_count = hc["_critical_aging"]
    need = len(joiners_30) + aging_critical_count - hc["_stock_ready"]
    lines.append(f"\n*🛒 Procurement Recommendation*")
    if need > 0:
        lines.append(f"• ⚠️ *Order {need} laptops*: {len(joiners_30)} joiners + {aging_critical_count} critical replacements vs {hc['_stock_ready']} in stock")
    else:
        lines.append(f"• ✅ Stock sufficient: {hc['_stock_ready']} available for {len(joiners_30)} joiners + {aging_critical_count} replacements")
    lines.append(f"• 90-day joiner forecast: {len(joiners_90)}")

    # ── 5. Top Critical Aging ──
    if critical_aging:
        lines.append(f"\n*🔝 Top Critical Aging* (showing 3 of {len(critical_aging)})")
        for a in critical_aging[:3]:
            lines.append(f"• {a['employee']} — {a['make']} {a['model']} ({a['age_years']}yr)")

    # ── 6. Upcoming Renewals ──
    if renewals:
        lines.append(f"\n*🔔 Upcoming Renewals*")
        for r in renewals[:5]:
            lines.append(f"• {r['app']} ({r['date'].strftime('%d %b')})")

    lines.append(f"\n_Full breakdown in the attached Word doc. React 👍/👎 or reply with feedback._")
    return "\n".join(lines)


def generate_monthly_full(data: dict, prev_snap: Optional[dict] = None) -> str:
    """Monthly full report includes everything from weekly + extra sections."""
    # Start with weekly full report content
    lines = [f"# IT Monthly Report — {TODAY.strftime('%B %Y')}\n"]

    # Include all weekly sections
    weekly = generate_weekly_full(data, prev_snap)
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
                lines.append(f"| {dept} | {model} | {qty} | {fmt_inr(price)} | {fmt_inr(total)} |")

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
    truly_assigned = sum(1 for r in data['assigned'] if is_truly_assigned(r))
    print(f"  Loaded: {truly_assigned} assigned laptops ({len(data['assigned'])} rows, "
          f"{len(data['assigned']) - truly_assigned} excluded as In Stock/blank), "
          f"{len(data['history'])} history records, {len(data['spend'])} spend rows, "
          f"{len(data['joinings'])} joiners")

    # Load previous snapshot for week-over-week comparison
    prev_snap = load_previous_snapshot()
    if prev_snap:
        print(f"  Previous snapshot: {prev_snap.get('date', 'unknown')} "
              f"(stock_ready={prev_snap.get('stock_ready')}, "
              f"stock_backup={prev_snap.get('stock_backup')})")
    else:
        print("  No previous snapshot found — this is the first run.")

    if report_type == "weekly":
        slack = generate_weekly_slack(data, prev_snap)
        full = generate_weekly_full(data, prev_snap)
    else:
        slack = generate_monthly_slack(data, prev_snap)
        full = generate_monthly_full(data, prev_snap)

    (OUTPUT_DIR / "slack-summary.md").write_text(slack, encoding="utf-8")
    (OUTPUT_DIR / "full-report.md").write_text(full, encoding="utf-8")

    # Dump structured metrics so the docx generator and chart generator can
    # produce visuals without re-parsing markdown.
    hc = get_health_check(data, prev_snap)
    laptop_spend = get_laptop_spend(data)
    app_total, _, _, _ = get_current_month_spend(data)
    runway = get_procurement_runway(data)
    pace = get_spend_pace(data)
    aging_dist = get_age_distribution(data)
    aging_all = get_aging_laptops(data)
    joiners_7 = get_upcoming_joiners(data, 7)
    joiners_30 = get_upcoming_joiners(data, 30)
    total_assigned = sum(1 for r in data["assigned"] if is_truly_assigned(r))

    overall = "🟢 On Track"
    if "🔴" in (hc["stock"], hc["aging"], hc["joiner_prep"], hc["spend"]):
        overall = "🔴 Action Needed"
    elif "🟡" in (hc["stock"], hc["aging"], hc["joiner_prep"], hc["spend"]):
        overall = "🟡 Attention"

    metrics = {
        "report_type": report_type,
        "date": TODAY.isoformat(),
        "overall_status": overall,
        "health": {k: hc[k] for k in ("stock", "aging", "joiner_prep", "spend")},
        "kpis": {
            "total_laptops": total_assigned + len(data["in_stock"]) + len(data["backup"]),
            "total_assigned": total_assigned,
            "stock_ready": len(data["in_stock"]),
            "stock_backup": len(data["backup"]),
            "aging_total": len(aging_all),
            "aging_critical": hc["_critical_aging"],
            "joiners_next_7": hc["_joiners_7"],
            "joiners_next_30": len(joiners_30),
            "runway_weeks": runway["weeks"],
            "avg_assignments_per_week": runway["avg_per_week"],
            "laptop_spend_month": laptop_spend["total_spend"],
            "laptop_spend_pct_of_budget": pace["pct_used"],
            "monthly_budget_inr": pace["monthly_planned"],
            "app_spend_month_usd": app_total,
        },
        "aging_distribution": aging_dist,
        "top_aging_critical": [
            {"name": a["employee"], "model": f"{a['make']} {a['model']}".strip(), "age_years": a["age_years"]}
            for a in aging_all[:10]
        ],
        "stock_vs_demand": {
            "stock_ready": len(data["in_stock"]),
            "stock_backup": len(data["backup"]),
            "joiners_7d": hc["_joiners_7"],
            "joiners_30d": len(joiners_30),
        },
        "risks": get_risk_callouts(data, hc),
    }
    (OUTPUT_DIR / "metrics.json").write_text(json.dumps(metrics, indent=2, default=str), encoding="utf-8")

    # Save new snapshot for next run
    save_snapshot(current_snapshot(data))
    print(f"  Saved new snapshot to {SNAPSHOT_DIR / 'latest.json'}")

    print(f"Reports saved to {OUTPUT_DIR}/")
    print(f"  slack-summary.md: {len(slack)} chars")
    print(f"  full-report.md: {len(full)} chars")


if __name__ == "__main__":
    main()
