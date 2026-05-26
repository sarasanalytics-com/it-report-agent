#!/usr/bin/env python3
"""
Generate PNG charts for the Word report from output/metrics.json.

Usage:
    python generate-charts.py

Reads:  output/metrics.json
Writes: output/charts/aging.png, stock.png, runway.png, kpi-strip.png
"""

import json
import pathlib
import sys

import matplotlib

matplotlib.use("Agg")  # headless
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

OUTPUT_DIR = pathlib.Path(__file__).resolve().parent.parent / "output"
CHARTS_DIR = OUTPUT_DIR / "charts"

# Brand palette — calm exec-friendly
NAVY = "#1F3A5F"
TEAL = "#2A9D8F"
AMBER = "#E9B949"
RED = "#E63946"
GREEN = "#43A047"
LIGHT_BG = "#F4F6FA"
INK = "#21314D"
MUTED = "#6B7A99"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.edgecolor": MUTED,
    "axes.labelcolor": INK,
    "xtick.color": INK,
    "ytick.color": INK,
    "axes.titlecolor": INK,
    "axes.titleweight": "bold",
    "axes.spines.top": False,
    "axes.spines.right": False,
})


def status_color(icon: str) -> str:
    return {"🟢": GREEN, "🟡": AMBER, "🔴": RED}.get(icon, MUTED)


def render_aging_donut(dist: dict, out: pathlib.Path) -> None:
    labels = list(dist.keys())
    values = [dist[k] for k in labels]
    if sum(values) == 0:
        return
    fig, ax = plt.subplots(figsize=(5.5, 4), dpi=140)
    fig.patch.set_facecolor("white")

    palette = [GREEN, TEAL, AMBER, "#F4A261", RED]
    palette = palette[:len(labels)]
    wedges, _ = ax.pie(
        values,
        labels=None,
        colors=palette,
        startangle=90,
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
    )
    # Centre stat
    total = sum(values)
    ax.text(0, 0.08, str(total), ha="center", va="center",
            fontsize=26, fontweight="bold", color=INK)
    ax.text(0, -0.18, "laptops", ha="center", va="center", fontsize=10, color=MUTED)

    ax.set_title("Laptop Age Distribution", fontsize=13, pad=14)
    ax.legend(wedges, [f"{l}: {v}" for l, v in zip(labels, values)],
              loc="center left", bbox_to_anchor=(1.0, 0.5),
              frameon=False, fontsize=10)
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_stock_vs_demand(svd: dict, out: pathlib.Path) -> None:
    fig, ax = plt.subplots(figsize=(6, 3.6), dpi=140)
    fig.patch.set_facecolor("white")

    categories = ["Ready stock", "Backup (3yr+)", "Joiners next 7d", "Joiners next 30d"]
    values = [svd["stock_ready"], svd["stock_backup"],
              svd["joiners_7d"], svd["joiners_30d"]]
    colors = [GREEN, TEAL, NAVY, MUTED]

    bars = ax.barh(categories, values, color=colors, height=0.6, edgecolor="white")
    for bar, v in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.02, bar.get_y() + bar.get_height() / 2,
                str(v), va="center", fontsize=11, fontweight="bold", color=INK)

    ax.invert_yaxis()
    ax.set_xlim(0, max(values) * 1.2 if max(values) > 0 else 5)
    ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    ax.set_title("Stock vs Joiner Demand", fontsize=13, pad=10, loc="left")
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_runway_gauge(weeks, out: pathlib.Path) -> None:
    weeks = weeks if weeks is not None else 0
    fig, ax = plt.subplots(figsize=(6, 2.6), dpi=140)
    fig.patch.set_facecolor("white")

    # Horizontal bar with zones
    ax.barh([0], [12], color="#EEF1F6", edgecolor="none", height=0.5)  # base track
    # Zones overlay
    ax.barh([0], [2], color="#FFE5E5", edgecolor="none", height=0.5)
    ax.barh([0], [2], left=2, color="#FFF4D6", edgecolor="none", height=0.5)
    ax.barh([0], [8], left=4, color="#E3F5E5", edgecolor="none", height=0.5)

    # Indicator
    color = RED if weeks < 2 else (AMBER if weeks < 4 else GREEN)
    ax.barh([0], [min(weeks, 12)], color=color, edgecolor="white", height=0.32)
    ax.text(min(weeks, 12) + 0.1, 0, f"  {weeks} weeks",
            va="center", fontsize=12, fontweight="bold", color=INK)

    # Zone labels
    for x, lbl, col in [(1, "Short", RED), (3, "Watch", AMBER), (8, "Healthy", GREEN)]:
        ax.text(x, -0.55, lbl, ha="center", va="top", fontsize=9, color=col, fontweight="bold")

    ax.set_xlim(0, 12)
    ax.set_ylim(-1, 0.6)
    ax.set_yticks([])
    ax.set_xticks([0, 2, 4, 8, 12])
    ax.set_xticklabels(["0", "2", "4", "8", "12+ wks"], fontsize=9, color=MUTED)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color(MUTED)
    ax.set_title("Procurement Runway", fontsize=13, pad=12, loc="left")
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _status_label(overall: str) -> tuple:
    """Strip emoji from overall status, return (label, color)."""
    if "🔴" in overall:
        return "ACTION NEEDED", RED
    if "🟡" in overall:
        return "ATTENTION", AMBER
    if "🟢" in overall:
        return "ON TRACK", GREEN
    return overall, MUTED


def render_kpi_strip(kpis: dict, health: dict, overall: str, date: str, out: pathlib.Path) -> None:
    """Hero banner — HR-focused KPI tiles (workforce, joiner pipeline, readiness, cost)."""
    fig, ax = plt.subplots(figsize=(11, 3.5), dpi=160)
    fig.patch.set_facecolor("white")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 42)
    ax.axis("off")

    # Header banner
    banner = FancyBboxPatch((0, 31), 100, 11, boxstyle="round,pad=0.02",
                            linewidth=0, facecolor=NAVY)
    ax.add_patch(banner)
    ax.text(2, 37.5, "Workforce & IT Readiness", color="white",
            fontsize=19, fontweight="bold", va="center")
    ax.text(2, 33.5, f"Week of {date}  ·  Saras Analytics", color="#C7D2E2",
            fontsize=10, va="center")

    # Status pill on the right
    label, color = _status_label(overall)
    pill_w = max(16, 1.4 * len(label))
    pill_x = 98 - pill_w
    ax.add_patch(FancyBboxPatch((pill_x, 34.5), pill_w, 4.5, boxstyle="round,pad=0.1",
                                linewidth=0, facecolor=color))
    ax.text(pill_x + pill_w / 2, 36.7, label, color="white",
            fontsize=10.5, fontweight="bold", ha="center", va="center")

    # 4 HR-focused KPI tiles
    headcount = kpis.get("total_assigned", 0)
    joiners_30 = kpis.get("joiners_next_30", 0)
    joiners_7 = kpis.get("joiners_next_7", 0)
    readiness = kpis.get("onboarding_readiness_pct")
    cost_per = kpis.get("cost_per_joiner_inr")

    readiness_color = MUTED
    if readiness is not None:
        readiness_color = GREEN if readiness >= 80 else (AMBER if readiness >= 50 else RED)
    joiner_color = status_color(health.get("joiner_prep", ""))

    fleet_size = kpis.get("total_laptops", 0)
    tiles = [
        ("Fleet Size", str(fleet_size), NAVY,
         f"{headcount} assigned · {kpis.get('stock_ready', 0)} ready"),
        ("Hiring Pipeline", str(joiners_30), TEAL,
         f"{joiners_7} in next 7 days"),
        ("Day-1 Readiness",
         f"{readiness:.0f}%" if readiness is not None else "—",
         readiness_color,
         f"of next-7d joiners ready" if readiness is not None else "no joiners next week"),
        ("Cost / Joiner",
         _short_inr(cost_per) if cost_per else "—",
         AMBER,
         "MTD laptop spend"),
    ]
    tile_w = 22
    gap = 2
    for i, (label, val, accent, sub) in enumerate(tiles):
        x = 2 + i * (tile_w + gap)
        ax.add_patch(FancyBboxPatch((x, 3), tile_w, 24, boxstyle="round,pad=0.02",
                                    linewidth=0, facecolor=LIGHT_BG))
        ax.add_patch(FancyBboxPatch((x, 3), 0.7, 24, boxstyle="round,pad=0",
                                    linewidth=0, facecolor=accent))
        ax.text(x + 1.8, 21.5, label.upper(), fontsize=8.5, color=MUTED,
                fontweight="bold", va="center")
        ax.text(x + 1.8, 13.5, val, fontsize=24, color=INK,
                fontweight="bold", va="center")
        if sub:
            ax.text(x + 1.8, 6, sub, fontsize=8.5, color=MUTED, va="center")

    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _short_inr(amount) -> str:
    try:
        n = float(amount)
    except (TypeError, ValueError):
        return "—"
    if n >= 1e7:
        return f"₹{n/1e7:.1f} Cr"
    if n >= 1e5:
        return f"₹{n/1e5:.1f} L"
    if n >= 1e3:
        return f"₹{n/1e3:.0f}k"
    return f"₹{n:.0f}"


def render_onboarding_pipeline(pipeline: dict, out: pathlib.Path) -> None:
    """Horizontal stacked timeline showing joiners over 7/14/30/90 day windows."""
    fig, ax = plt.subplots(figsize=(9, 2.8), dpi=160)
    fig.patch.set_facecolor("white")

    stages = [
        ("Next 7 days",  pipeline.get("d7", 0),  RED),
        ("Next 14 days", pipeline.get("d14", 0), AMBER),
        ("Next 30 days", pipeline.get("d30", 0), TEAL),
        ("Next 90 days", pipeline.get("d90", 0), NAVY),
    ]
    labels = [s[0] for s in stages]
    values = [s[1] for s in stages]
    colors = [s[2] for s in stages]

    bars = ax.bar(labels, values, color=colors, width=0.55, edgecolor="white", linewidth=2)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + max(values) * 0.03 + 0.2,
                str(v), ha="center", va="bottom",
                fontsize=14, fontweight="bold", color=INK)

    ax.set_ylim(0, max(values) * 1.25 if max(values) > 0 else 5)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color(MUTED)
    ax.tick_params(axis="x", labelsize=10, colors=INK)
    ax.set_title("Hiring Pipeline", fontsize=13, pad=12, loc="left", color=INK)
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_spend_progress(spend_pct, monthly_spend_inr, monthly_budget_inr, out: pathlib.Path) -> None:
    """Horizontal progress bar of MTD laptop spend vs monthly budget."""
    fig, ax = plt.subplots(figsize=(9, 2.2), dpi=160)
    fig.patch.set_facecolor("white")

    pct = float(spend_pct) if spend_pct is not None else 0
    pct_capped = min(pct, 120)

    # Background track
    ax.barh([0], [100], color="#EEF1F6", edgecolor="none", height=0.45)
    # Zones overlay (>100% = red)
    if pct > 100:
        ax.barh([0], [pct_capped - 100], left=100, color="#FBE5E7", edgecolor="none", height=0.45)

    color = GREEN if pct < 90 else (AMBER if pct < 110 else RED)
    ax.barh([0], [pct_capped], color=color, edgecolor="white", height=0.30)

    # Big label
    ax.text(pct_capped + 1, 0, f"  {pct:.0f}%",
            va="center", fontsize=13, fontweight="bold", color=INK)
    # Sub-label
    spent = _short_inr(monthly_spend_inr) if monthly_spend_inr else "—"
    budget = _short_inr(monthly_budget_inr) if monthly_budget_inr else "—"
    ax.text(0, -0.7, f"{spent} of {budget} monthly budget",
            fontsize=9.5, color=MUTED, va="top")

    ax.set_xlim(0, 125)
    ax.set_ylim(-1.2, 0.6)
    ax.set_yticks([])
    ax.set_xticks([0, 50, 100, 120])
    ax.set_xticklabels(["0%", "50%", "100%", "120%"], fontsize=9, color=MUTED)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color(MUTED)
    ax.set_title("Laptop Spend Pace (MTD vs Monthly Budget)",
                 fontsize=13, pad=12, loc="left", color=INK)
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    metrics_path = OUTPUT_DIR / "metrics.json"
    if not metrics_path.exists():
        print(f"metrics.json not found at {metrics_path}", file=sys.stderr)
        sys.exit(1)
    metrics = json.loads(metrics_path.read_text(encoding="utf-8"))
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Generating charts in {CHARTS_DIR} …")

    render_kpi_strip(
        metrics["kpis"], metrics["health"], metrics["overall_status"],
        metrics["date"], CHARTS_DIR / "kpi-strip.png",
    )
    print("  [ok]kpi-strip.png")

    render_aging_donut(metrics["aging_distribution"], CHARTS_DIR / "aging.png")
    print("  [ok]aging.png")

    render_stock_vs_demand(metrics["stock_vs_demand"], CHARTS_DIR / "stock.png")
    print("  [ok]stock.png")

    render_runway_gauge(metrics["kpis"].get("runway_weeks"), CHARTS_DIR / "runway.png")
    print("  [ok]runway.png")

    if metrics.get("onboarding_pipeline"):
        render_onboarding_pipeline(metrics["onboarding_pipeline"],
                                   CHARTS_DIR / "pipeline.png")
        print("  [ok]pipeline.png")

    kpis = metrics.get("kpis", {})
    if kpis.get("laptop_spend_pct_of_budget") is not None or kpis.get("laptop_spend_month"):
        render_spend_progress(kpis.get("laptop_spend_pct_of_budget"),
                              kpis.get("laptop_spend_month"),
                              kpis.get("monthly_budget_inr"),
                              CHARTS_DIR / "spend.png")
        print("  [ok]spend.png")


if __name__ == "__main__":
    main()
