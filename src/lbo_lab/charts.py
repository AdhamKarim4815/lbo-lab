"""Matplotlib charts for the README (the workbook has native Excel charts of its own)."""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

from .engine import ModelResult  # noqa: E402
from .sensitivity import Grid  # noqa: E402

NAVY, STEEL, LIGHT_BLUE = "#1F3864", "#2E75B6", "#8FAADC"
GREEN, RED, AMBER, GREY = "#3C8D5A", "#C0392B", "#E9A23B", "#6B6B6B"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10, "axes.unicode_minus": False,
    "axes.spines.top": False, "axes.spines.right": False, "axes.edgecolor": "#999999",
    "axes.titleweight": "bold", "axes.titlesize": 13, "axes.titlelocation": "left",
    "figure.dpi": 100, "savefig.dpi": 170, "savefig.bbox": "tight",
})


def _subtitle(res: ModelResult) -> str:
    r = res.returns
    return (f"{res.assumptions.case.name} case, exit year {r.exit_year} at {r.exit_multiple:.1f}x  |  "
            f"sponsor IRR {100 * r.irr:.1f}%, MOIC {r.moic:.2f}x  |  Halcyon Packaging Co. (fictional)")


def _labels(res: ModelResult) -> list[str]:
    return ["Close"] + [f"FY{res.assumptions.base_year + t}E" for t in range(1, res.n + 1)]


def bridge_chart(res: ModelResult, path: Path) -> Path:
    b = res.returns.bridge
    steps = [("Entry\nequity", b["entry_equity"], True), ("EBITDA\ngrowth", b["ebitda_growth"], False),
             ("Multiple\nexpansion", b["multiple_expansion"], False),
             ("Net debt\npaydown", b["net_debt_paydown"], False),
             ("Transaction\ncosts", b["transaction_costs"], False), ("Exit\nequity", b["exit_equity"], True)]
    fig, ax = plt.subplots(figsize=(9, 5.2))
    level = 0.0
    for i, (_, val, total) in enumerate(steps):
        if total:
            ax.bar(i, val, color=NAVY, width=0.6)
            level = val
            ax.text(i, val + 12, f"{val:,.0f}", ha="center", va="bottom", fontweight="bold")
        else:
            lo = min(level, level + val)
            ax.bar(i, abs(val), bottom=lo, color=GREEN if val >= 0 else RED, width=0.6)
            ax.text(i, max(level, level + val) + 12, f"{val:+,.0f}" if round(val) else "0", ha="center", va="bottom",
                    fontweight="bold", color=GREEN if val >= 0 else RED)
            level += val
        if i < len(steps) - 1:
            ax.plot([i + 0.3, i + 0.7], [level, level], color="#999999", lw=0.8)
    ax.set_xticks(range(len(steps)), [s[0] for s in steps])
    ax.set_ylabel("USD millions (100% of equity)")
    ax.yaxis.grid(True, color="#E3E3E3")
    ax.set_axisbelow(True)
    ax.set_title("Value creation bridge", pad=22)
    ax.set_ylim(0, max(b["exit_equity"], b["entry_equity"] + b["ebitda_growth"] + b["multiple_expansion"]
                       + b["net_debt_paydown"]) * 1.12)
    ax.text(0, 1.015, _subtitle(res), transform=ax.transAxes, fontsize=8.5, color=GREY)
    fig.savefig(path)
    plt.close(fig)
    return path


def credit_chart(res: ModelResult, path: Path) -> Path:
    labels, n = _labels(res), res.n
    xs = list(range(n + 1))
    c = res.credit
    fig, ax = plt.subplots(figsize=(9, 5))
    w = 0.38
    ax.bar([x - w / 2 for x in xs], c["total_leverage"], w, color=NAVY, label="Total debt / EBITDA")
    ax.bar([x + w / 2 for x in xs], c["senior_leverage"], w, color=LIGHT_BLUE, label="Senior debt / EBITDA")
    ax.set_xticks(xs, labels)
    ax.set_ylabel("Leverage (x)")
    ax.yaxis.grid(True, color="#E3E3E3")
    ax.set_axisbelow(True)
    ax2 = ax.twinx()
    ax2.spines["right"].set_visible(True)
    cov = c["interest_coverage"][1:]
    ax2.plot(xs[1:], cov, color=AMBER, marker="o", lw=2.2, label="Interest coverage (right axis)")
    fccr = c["fccr"][1:]
    ax2.plot(xs[1:], fccr, color=GREY, marker="s", lw=1.5, ls="--", label="FCCR (right axis)")
    ax2.set_ylabel("Coverage (x)")
    ax2.set_ylim(0, max(max(cov), max(fccr)) * 1.25)
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=4, frameon=False)
    ax.set_title("Leverage and coverage", pad=22)
    ax.text(0, 1.015, _subtitle(res), transform=ax.transAxes, fontsize=8.5, color=GREY)
    fig.savefig(path)
    plt.close(fig)
    return path


def debt_chart(res: ModelResult, path: Path) -> Path:
    labels, n = _labels(res), res.n
    xs = list(range(n + 1))
    rows = res.rows
    fig, ax = plt.subplots(figsize=(9, 5))
    bottom = [0.0] * (n + 1)
    for key, name, color in (("tl_close", "Senior term loan", NAVY), ("notes_close", "Subordinated notes", LIGHT_BLUE),
                             ("rev_close", "Revolver", AMBER)):
        ax.bar(xs, rows[key], 0.6, bottom=bottom, color=color, label=name)
        bottom = [b + v for b, v in zip(bottom, rows[key], strict=True)]
    for x, tot in zip(xs, bottom, strict=True):
        ax.text(x, tot + 8, f"{tot:,.0f}", ha="center", fontsize=9, fontweight="bold")
    ax.plot(xs, rows["cash_close"], color=GREEN, marker="o", lw=1.8, label="Cash")
    ax.set_xticks(xs, labels)
    ax.set_ylabel("USD millions")
    ax.yaxis.grid(True, color="#E3E3E3")
    ax.set_axisbelow(True)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=4, frameon=False)
    ax.set_title("Debt paydown by tranche", pad=22)
    ax.set_ylim(0, max(bottom) * 1.12)
    ax.text(0, 1.015, _subtitle(res), transform=ax.transAxes, fontsize=8.5, color=GREY)
    fig.savefig(path)
    plt.close(fig)
    return path


def sensitivity_chart(grid: Grid, path: Path, subtitle: str = "") -> Path:
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    data = [[100 * v for v in row] for row in grid.irr]
    lo, hi = min(map(min, data)), max(map(max, data))
    im = ax.imshow(data, cmap="RdYlGn", aspect="auto", vmin=lo, vmax=hi)
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            centre = i == len(data) // 2 and j == len(row) // 2
            share = (v - lo) / (hi - lo) if hi > lo else 0.5
            ax.text(j, i, f"{v:.1f}%", ha="center", va="center", fontsize=9,
                    color="white" if share < 0.12 or share > 0.92 else "black",
                    fontweight="bold" if centre else "normal")
    ax.set_xticks(range(len(grid.col_values)), [f"{v:.1f}x" for v in grid.col_values])
    ax.set_yticks(range(len(grid.row_values)), [f"{v:.1f}x" for v in grid.row_values])
    ax.set_xlabel(grid.col_label)
    ax.set_ylabel(grid.row_label)
    ax.spines[:].set_visible(False)
    ax.set_title("Sponsor IRR sensitivity", pad=22)
    ax.text(0, 1.02, subtitle or grid.name, transform=ax.transAxes, fontsize=8.5, color=GREY)
    fig.colorbar(im, ax=ax, format="%.0f%%", shrink=0.85)
    fig.savefig(path)
    plt.close(fig)
    return path


def make_all(res: ModelResult, grid: Grid, outdir: Path) -> list[Path]:
    outdir.mkdir(parents=True, exist_ok=True)
    sub = f"Entry vs exit EV/EBITDA, {res.assumptions.case.name} case, exit year {res.returns.exit_year}"
    return [bridge_chart(res, outdir / "value_creation_bridge.png"),
            credit_chart(res, outdir / "leverage_coverage.png"),
            debt_chart(res, outdir / "debt_paydown.png"),
            sensitivity_chart(grid, outdir / "sensitivity_irr.png", sub)]
