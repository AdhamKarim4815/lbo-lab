from __future__ import annotations

import pytest

from lbo_lab import sensitivity
from lbo_lab.assumptions import SCENARIO_KEYS, Assumptions
from lbo_lab.engine import run_model


@pytest.fixture(scope="module")
def grids():
    a = Assumptions()
    return a, sensitivity.entry_vs_exit(a), sensitivity.leverage_vs_exit_year(a), \
        sensitivity.exit_year_vs_exit_multiple(a)


@pytest.mark.parametrize("key", SCENARIO_KEYS)
def test_higher_entry_multiple_lowers_irr_and_moic(key):
    g = sensitivity.entry_vs_exit(Assumptions().with_scenario(key))
    for grid in (g.irr, g.moic):
        for j in range(5):
            col = [grid[i][j] for i in range(5)]
            assert all(x > y for x, y in zip(col, col[1:], strict=False))


@pytest.mark.parametrize("key", SCENARIO_KEYS)
def test_higher_exit_multiple_raises_irr_and_moic(key):
    a = Assumptions().with_scenario(key)
    for g in (sensitivity.entry_vs_exit(a), sensitivity.exit_year_vs_exit_multiple(a)):
        for data in (g.irr, g.moic):
            if g.col_label == "Exit EV/EBITDA":
                for row in data:
                    assert all(y > x for x, y in zip(row, row[1:], strict=False))
            else:   # rows are exit multiples
                for j in range(len(data[0])):
                    col = [data[i][j] for i in range(len(data))]
                    assert all(y > x for x, y in zip(col, col[1:], strict=False))


def test_grid_shapes_and_axes(grids):
    a, g1, g2, g3 = grids
    assert (len(g1.irr), len(g1.irr[0])) == (5, 5)
    assert (len(g2.irr), len(g2.irr[0])) == (5, 5)
    assert g1.row_values == [8.0, 8.5, 9.0, 9.5, 10.0]
    assert g3.col_values == [3.0, 4.0, 5.0, 6.0, 7.0]
    assert g2.row_values[2] == pytest.approx(a.term_loan_x + a.sub_notes_x)


def test_centres_equal_the_base_run(grids):
    a, g1, g2, g3 = grids
    r = run_model(a).returns
    assert g1.irr[2][2] == pytest.approx(r.irr, abs=1e-12)
    assert g1.moic[2][2] == pytest.approx(r.moic, abs=1e-12)
    assert g2.irr[2][2] == pytest.approx(r.irr, abs=1e-12)   # base leverage, exit year 5
    assert g3.irr[2][2] == pytest.approx(r.irr, abs=1e-12)


def test_grid_cell_equals_a_direct_run(grids):
    a, g1, _, g3 = grids
    direct = run_model(a, entry_multiple=9.5, exit_multiple=8.5).returns
    assert g1.irr[3][1] == pytest.approx(direct.irr, abs=1e-12)
    direct = run_model(a, exit_multiple=9.5, exit_year=7).returns
    assert g3.moic[3][4] == pytest.approx(direct.moic, abs=1e-12)


def test_leverage_changes_the_debt_schedule(grids):
    a, _, g2, _ = grids
    low = run_model(a, term_loan_x=a.term_loan_x * 4.5 / 5.5, sub_notes_x=a.sub_notes_x * 4.5 / 5.5)
    high = run_model(a, term_loan_x=a.term_loan_x * 6.5 / 5.5, sub_notes_x=a.sub_notes_x * 6.5 / 5.5)
    assert high.su["total_debt"] > low.su["total_debt"]
    assert g2.irr[4][2] > g2.irr[0][2]   # more leverage lifts base-case returns


def test_entry_multiple_does_not_change_the_debt_schedule():
    """This is what lets the Excel grid use closed-form formulas."""
    a = Assumptions()
    r1, r2 = run_model(a), run_model(a, entry_multiple=10.5)
    assert r1.rows["total_debt"] == pytest.approx(r2.rows["total_debt"])
    assert r1.rows["cash_close"] == pytest.approx(r2.rows["cash_close"])
