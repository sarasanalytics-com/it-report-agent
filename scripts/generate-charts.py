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


def render_kpi_strip(kpis: dict, health: dict, overall: str, date: str, out: pathlib.Path) -> None:
    """Hero banner image: title + status + 4 KPI tiles."""
    fig, ax = plt.subplots(figsize=(11, 3.3), dpi=160)
    fig.patch.set_facecolor("white")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 40)
    ax.axis("off")

    # Header banner
    banner = FancyBboxPatch((0, 30), 100, 10, boxstyle="round,pad=0.02",
                            linewidth=0, facecolor=NAVY)
    ax.add_patch(banner)
    ax.text(2, 36.2, "IT Operations Report", color="white",
            fontsize=18, fontweight="bold", va="center")
    ax.text(2, 32.4, date, color="#C7D2E2", fontsize=10, va="center")
    ax.text(98, 35, overall, color="white", fontsize=12,
            fontweight="bold", ha="right", va="center")

    # 4 KPI tiles
    tiles = [
        ("Total Laptops", str(kpis["total_laptops"]), MUTED, ""),
        ("Stock Ready", str(kpis["stock_ready"]), status_color(health["stock"]), ""),
        ("Critical Aging", str(kpis["aging_critical"]), status_color(health["aging"]), f"{kpis['aging_total']} total >3.5y"),
        ("Joiners next 30d", str(kpis["joiners_next_30"]), status_color(health["joiner_prep"]),
         f"runway {kpis['runway_weeks']}w" if kpis.get("runway_weeks") is not None else ""),
    ]
    tile_w = 22
    gap = 2
    for i, (label, val, accent, sub) in enumerate(tiles):
        x = 2 + i * (tile_w + gap)
        # Card body
        ax.add_patch(FancyBboxPatch((x, 4), tile_w, 22, boxstyle="round,pad=0.02",
                                    linewidth=0, facecolor=LIGHT_BG))
        # Accent stripe
        ax.add_patch(FancyBboxPatch((x, 4), 0.6, 22, boxstyle="round,pad=0",
                                    linewidth=0, facecolor=accent))
        ax.text(x + 1.8, 20, label, fontsize=9, color=MUTED,
                fontweight="bold", va="center")
        ax.text(x + 1.8, 13, val, fontsize=22, color=INK,
                fontweight="bold", va="center")
        if sub:
            ax.text(x + 1.8, 7, sub, fontsize=8, color=MUTED, va="center")

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
    print("  ✓ kpi-strip.png")

    render_aging_donut(metrics["aging_distribution"], CHARTS_DIR / "aging.png")
    print("  ✓ aging.png")

    render_stock_vs_demand(metrics["stock_vs_demand"], CHARTS_DIR / "stock.png")
    print("  ✓ stock.png")

    render_runway_gauge(metrics["kpis"].get("runway_weeks"), CHARTS_DIR / "runway.png")
    print("  ✓ runway.png")


if __name__ == "__main__":
    main()
