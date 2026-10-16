"""Shared fixtures. Counts every Excel-vs-Python comparison for the terminal summary."""

from __future__ import annotations

import math
import shutil
from dataclasses import replace
from pathlib import Path

import openpyxl
import pytest

from lbo_lab import excel, recalc
from lbo_lab.assumptions import Assumptions

RECON = {"checks": 0}

# name -> assumption overrides. "stress" draws the revolver; "pik" accrues the notes;
# "breach" is over-levered on purpose so the workbook must flag a minimum cash breach.
VARIANTS: dict[str, dict] = {
    "base": {"scenario": "base"},
    "management": {"scenario": "management"},
    "downside": {"scenario": "downside"},
    "stress": {"scenario": "downside", "revolver_draw_at_close": 10.0, "sweep_pct": 1.0,
               "rollover_equity": 0.0, "option_pool_pct": 0.10, "exit_year": 6},
    "pik": {"scenario": "management", "notes_pik": True, "exit_year": 7, "rollover_equity": 40.0,
            "option_pool_pct": 0.0},
    "breach": {"scenario": "downside", "term_loan_x": 4.75, "sub_notes_x": 1.75,
               "mandatory_amort": (0.05,) * 7},
}
CLEAN = [n for n in VARIANTS if n != "breach"]


def assumptions_for(name: str) -> Assumptions:
    return replace(Assumptions(), **VARIANTS[name])


def close(actual, expected, label: str, tol: float = 1e-6) -> None:
    """Compare an Excel value with the engine value; NaN in Python must be 'n.a.' in Excel."""
    RECON["checks"] += 1
    if isinstance(expected, float) and math.isnan(expected):
        assert actual in ("n.a.", None), f"{label}: expected n.a., got {actual!r}"
        return
    assert isinstance(actual, (int, float)), f"{label}: Excel returned {actual!r}"
    assert abs(actual - expected) <= tol * max(1.0, abs(expected)), f"{label}: excel {actual} vs python {expected}"


@pytest.fixture(scope="session")
def soffice():
    if recalc.find_soffice() is None:
        pytest.skip("LibreOffice (soffice) is not installed")


@pytest.fixture(scope="session")
def recalculated(tmp_path_factory, soffice):
    """Build each variant, recalculate it headlessly in LibreOffice, and read the results."""
    cache: dict[str, tuple] = {}

    def get(name: str):
        if name not in cache:
            base = tmp_path_factory.mktemp(name)
            a = assumptions_for(name)
            src = base / "model.xlsx"
            layout = excel.build_workbook(src, a)
            out = recalc.recalculate(src, base / "recalc")
            wb = openpyxl.load_workbook(out, data_only=True)
            cache[name] = (a, layout, wb)
        return cache[name]

    return get


def pytest_terminal_summary(terminalreporter):
    if RECON["checks"]:
        terminalreporter.write_line(f"Excel vs Python reconciliation checks executed: {RECON['checks']}")


@pytest.fixture(scope="session")
def have_soffice() -> bool:
    return shutil.which("soffice") is not None


@pytest.fixture
def tmp_out(tmp_path) -> Path:
    return tmp_path
