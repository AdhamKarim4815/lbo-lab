"""Build the workbook, recalculate it in headless LibreOffice, and reconcile with the Python engine.

Skipped automatically when LibreOffice is not installed.
"""

from __future__ import annotations

import pytest

from conftest import CLEAN, VARIANTS, close
from lbo_lab import sensitivity
from lbo_lab.engine import run_model

NAMES = list(VARIANTS)  # includes the deliberately breaching variant
YC = "DEFGHIJK"


def _row(layout, sheet, key):
    return layout.rows[sheet][key]


@pytest.mark.parametrize("name", NAMES)
def test_sources_and_uses(recalculated, name):
    a, layout, wb = recalculated(name)
    su = run_model(a).su
    ws = wb["Transaction"]
    mapping = {"s_rev": "revolver", "s_tl": "term_loan", "s_notes": "sub_notes", "s_roll": "rollover",
               "s_sponsor": "sponsor_equity", "s_total": "total_sources", "u_ev": "enterprise_value",
               "u_txn": "transaction_fees", "u_fin": "financing_fees", "u_cash": "min_cash",
               "u_total": "total_uses"}
    for key, py in mapping.items():
        close(ws[f"D{_row(layout, 'Transaction', key)}"].value, su[py], f"{name} {key}")
    close(ws[f"D{_row(layout, 'Transaction', 'su_check')}"].value, 0.0, f"{name} S&U check")


@pytest.mark.parametrize("name", NAMES)
def test_operating_and_fcf(recalculated, name):
    a, layout, wb = recalculated(name)
    res = run_model(a)
    ws = wb["Operating"]
    keys = {"revenue": "revenue", "ebitda": "ebitda", "ebit": "ebit", "capex": "capex", "dnwc": "dnwc",
            "tax": "tax", "ni": "net_income", "cfads": "cfads", "fcf_post_int": "fcf_post_int",
            "fcf_post_mand": "fcf_post_mand", "ebt": "ebt"}
    for key, py in keys.items():
        for t in range(1, 8):
            close(ws[f"{YC[t]}{_row(layout, 'Operating', key)}"].value, res.rows[py][t], f"{name} {key} t{t}")
    for t in (0, 1, 7):
        close(ws[f"{YC[t]}{_row(layout, 'Operating', 'ebitda')}"].value, res.rows["ebitda"][t], f"{name} ebitda")


@pytest.mark.parametrize("name", NAMES)
def test_debt_schedules_and_credit(recalculated, name):
    a, layout, wb = recalculated(name)
    res = run_model(a)
    ws = wb["Debt"]
    keys = {"rev_close": "rev_close", "tl_close": "tl_close", "notes_close": "notes_close",
            "cash_close": "cash_close", "total_debt": "total_debt", "net_debt": "net_debt",
            "lev_total": ("total_leverage", 0), "lev_senior": ("senior_leverage", 0),
            "lev_net": ("net_leverage", 0)}
    for key, py in keys.items():
        for t in range(0, 8):
            expected = res.rows[py][t] if isinstance(py, str) else res.credit[py[0]][t]
            close(ws[f"{YC[t]}{_row(layout, 'Debt', key)}"].value, expected, f"{name} {key} t{t}")
    flows = {"rev_draw": "rev_draw", "rev_repay": "rev_repay", "tl_mand": "mand", "tl_sweep": "sweep",
             "notes_pik": "int_notes_pik", "gross_int": "gross_cash_interest", "net_cash_int": "net_cash_interest",
             "int_inc": "int_income"}
    for key, py in flows.items():
        for t in range(1, 8):
            close(ws[f"{YC[t]}{_row(layout, 'Debt', key)}"].value, res.rows[py][t], f"{name} {key} t{t}")
    for key, py in (("cov_ebitda", "interest_coverage"), ("cov_capex", "capex_coverage"), ("fccr", "fccr")):
        for t in range(1, 8):
            close(ws[f"{YC[t]}{_row(layout, 'Debt', key)}"].value, res.credit[py][t], f"{name} {key} t{t}")


@pytest.mark.parametrize("name", NAMES)
def test_exit_returns_and_bridge(recalculated, name):
    a, layout, wb = recalculated(name)
    r = run_model(a).returns
    ws = wb["Returns"]

    def cell(key, col="D"):
        return ws[f"{col}{_row(layout, 'Returns', key)}"].value

    pairs = {"x_ebitda": r.exit_ebitda, "x_ev": r.exit_ev, "x_nd": r.exit_net_debt, "x_equity": r.exit_equity,
             "d_pool": r.option_pool_value, "d_sponsor": r.sponsor_proceeds, "moic": r.moic, "irr": r.irr,
             "irr_check": r.irr_from_flows, "b_e0": r.bridge["entry_equity"],
             "b_growth": r.bridge["ebitda_growth"], "b_mult": r.bridge["multiple_expansion"],
             "b_debt": r.bridge["net_debt_paydown"], "b_fees": r.bridge["transaction_costs"],
             "b_gain": r.bridge["total_gain"], "b_exit": r.bridge["exit_equity"],
             "b_pool": r.sponsor_bridge["option_pool_dilution"], "b_sgain": r.sponsor_bridge["sponsor_gain"],
             "b_check": 0.0, "b_scheck": 0.0}
    for key, expected in pairs.items():
        close(cell(key), expected, f"{name} {key}", tol=1e-6)


@pytest.mark.parametrize("name", NAMES)
def test_sensitivity_grids(recalculated, name):
    a, layout, wb = recalculated(name)
    ws = wb["Sensitivity"]
    res = run_model(a)
    g1 = sensitivity.entry_vs_exit(a)
    g3 = sensitivity.exit_year_vs_exit_multiple(a)
    for grid, key in ((g1, "entry_exit"), (g3, "exit_year_exit_mult")):
        for metric, data in (("moic", grid.moic), ("irr", grid.irr)):
            spec = layout.grids[key][metric]
            for i, row in enumerate(data):
                for j, expected in enumerate(row):
                    cell = f"{spec['cols'][j]}{spec['first_row'] + i}"
                    close(ws[cell].value, expected, f"{name} {key} {metric} [{i},{j}]")
        first = layout.grids[key]["moic"]["first_row"]
        for i, rv in enumerate(grid.row_values):
            close(ws[f"C{first + i}"].value, rv, f"{name} {key} row axis {i}")
    close(g1.irr[2][2], res.returns.irr, f"{name} grid centre equals model IRR", tol=1e-12)


@pytest.mark.parametrize("name", CLEAN)
def test_checks_sheet_all_ok(recalculated, name):
    a, layout, wb = recalculated(name)
    ws = wb["Checks"]
    rows = layout.rows["Checks"]
    statuses = [ws[f"F{rows[f'c{i}']}"].value for i in range(1, 21)]
    assert statuses == ["OK"] * 20, statuses
    assert ws[f"F{rows['master']}"].value == "ALL CHECKS PASS"
    assert ws[f"D{rows['master']}"].value == 0


def test_scenario_dropdown_switches_case(recalculated):
    """The three base-sheet builds differ only in the dropdown cell, and Excel follows it."""
    irrs = []
    for name in ("base", "management", "downside"):
        a, layout, wb = recalculated(name)
        assert wb["Inputs"]["D5"].value == a.case.name
        irrs.append(wb["Returns"][f"D{_row(layout, 'Returns', 'irr')}"].value)
    assert irrs[1] > irrs[0] > irrs[2]


def test_stress_case_exercises_revolver(recalculated):
    a, layout, wb = recalculated("stress")
    res = run_model(a)
    assert max(res.rows["rev_draw"]) > 0 and max(res.rows["rev_repay"]) > 0, "stress should draw and repay"
    ws = wb["Debt"]
    close(ws[f"K{_row(layout, 'Debt', 'rev_close')}"].value, res.rows["rev_close"][7], "stress revolver")


def test_pik_case_accrues_notes(recalculated):
    a, layout, wb = recalculated("pik")
    res = run_model(a)
    assert res.rows["notes_close"][7] > res.rows["notes_close"][0]
    close(wb["Debt"][f"K{_row(layout, 'Debt', 'notes_close')}"].value, 150.0 * 1.12**7, "PIK compounding", 1e-9)


def test_breach_variant_is_flagged_by_the_checks_sheet(recalculated):
    a, layout, wb = recalculated("breach")
    ws = wb["Checks"]
    rows = layout.rows["Checks"]
    assert min(run_model(a).rows["cash_close"]) < a.min_cash
    assert ws[f"F{rows['c4']}"].value == "ERROR", "cash below minimum must be flagged"
    assert ws[f"F{rows['master']}"].value == "CHECK FAILURES"
