"""Sensitivity grids computed by re-running the full engine for every cell.

The Excel workbook rebuilds the first and third grids with closed-form live formulas
(valid because entry multiple, exit multiple and exit year do not change the debt
schedule). The leverage grid needs the debt schedule re-solved, so it is Python only.
"""

from __future__ import annotations

from dataclasses import dataclass

from .assumptions import Assumptions
from .engine import run_model

STEPS = (-2, -1, 0, 1, 2)
MULTIPLE_STEP = 0.5
LEVERAGE_STEP = 0.5
EXIT_YEARS = (3, 4, 5, 6, 7)


@dataclass
class Grid:
    name: str
    row_label: str
    col_label: str
    row_values: list[float]
    col_values: list[float]
    irr: list[list[float]]
    moic: list[list[float]]


def _axis(center: float, step: float) -> list[float]:
    return [center + k * step for k in STEPS]


def _fill(a: Assumptions, name: str, row_label: str, col_label: str,
          rows: list[float], cols: list[float], make) -> Grid:
    irr_g: list[list[float]] = []
    moic_g: list[list[float]] = []
    for r in rows:
        irr_row, moic_row = [], []
        for c in cols:
            res = run_model(a, **make(r, c))
            irr_row.append(res.returns.irr)
            moic_row.append(res.returns.moic)
        irr_g.append(irr_row)
        moic_g.append(moic_row)
    return Grid(name, row_label, col_label, list(rows), list(cols), irr_g, moic_g)


def entry_vs_exit(a: Assumptions, step: float = MULTIPLE_STEP) -> Grid:
    """Rows: entry EV/EBITDA. Columns: exit EV/EBITDA. Exit year as in ``a``."""
    rows = _axis(a.entry_multiple, step)
    cols = _axis(a.effective_exit_multiple, step)
    return _fill(a, "Entry multiple vs exit multiple", "Entry EV/EBITDA", "Exit EV/EBITDA",
                 rows, cols, lambda r, c: {"entry_multiple": r, "exit_multiple": c})


def leverage_vs_exit_year(a: Assumptions, step: float = LEVERAGE_STEP) -> Grid:
    """Rows: total debt / LTM EBITDA at close (term loan and notes scaled pro rata)."""
    base = a.term_loan_x + a.sub_notes_x
    rows = _axis(base, step)
    cols = [float(y) for y in EXIT_YEARS if y <= a.horizon]

    def make(lev: float, year: float) -> dict[str, object]:
        k = lev / base
        return {"term_loan_x": a.term_loan_x * k, "sub_notes_x": a.sub_notes_x * k,
                "exit_year": int(year)}

    return _fill(a, "Leverage vs exit year", "Total debt / EBITDA", "Exit year", rows, cols, make)


def exit_year_vs_exit_multiple(a: Assumptions, step: float = MULTIPLE_STEP) -> Grid:
    """Rows: exit EV/EBITDA. Columns: exit year."""
    rows = _axis(a.effective_exit_multiple, step)
    cols = [float(y) for y in EXIT_YEARS if y <= a.horizon]
    return _fill(a, "Exit year vs exit multiple", "Exit EV/EBITDA", "Exit year", rows, cols,
                 lambda r, c: {"exit_multiple": r, "exit_year": int(c)})


def all_grids(a: Assumptions) -> list[Grid]:
    return [entry_vs_exit(a), leverage_vs_exit_year(a), exit_year_vs_exit_multiple(a)]
