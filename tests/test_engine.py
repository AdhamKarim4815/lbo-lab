"""Engine unit tests against hand-computed cases."""

from __future__ import annotations

import math
from dataclasses import replace

import pytest

from helpers import simple, with_case
from lbo_lab.assumptions import SCENARIO_KEYS, Assumptions, resolve_scenario
from lbo_lab.engine import bridge_residual, irr, irr_from_moic, run_model


# ------------------------------------------------------------------------------ IRR
def test_irr_two_cash_flows():
    assert irr([-100, 0, 0, 0, 0, 200]) == pytest.approx(2 ** 0.2 - 1, abs=1e-12)


def test_irr_with_interim_flows():
    assert irr([-100, 10, 10, 110]) == pytest.approx(0.10, abs=1e-10)


def test_irr_negative_return():
    assert irr([-100, 0, 50]) == pytest.approx(math.sqrt(0.5) - 1, abs=1e-10)


def test_irr_needs_both_signs():
    assert math.isnan(irr([100, 100]))
    assert math.isnan(irr([-100, -5]))


def test_irr_from_moic_edges():
    assert irr_from_moic(2.0, 1) == pytest.approx(1.0)
    assert irr_from_moic(0.0, 5) == -1.0
    assert irr_from_moic(-1.0, 5) == -1.0


# ------------------------------------------------------------- hand-computed returns
def test_buy_at_8x_no_paydown_cash_builds():
    """EBITDA 100 for five years, no debt, no tax: equity 800 returns 800 EV + 500 cash."""
    res = run_model(simple())
    r = res.returns
    assert r.exit_equity == pytest.approx(1300.0)
    assert r.moic == pytest.approx(1.625)
    assert r.irr == pytest.approx(1.625 ** 0.2 - 1)
    assert r.irr_from_flows == pytest.approx(r.irr, abs=1e-10)


def test_irr_equals_ebitda_growth_when_fcf_is_zero_and_multiple_flat():
    """Capex equal to EBITDA means no cash builds; EV grows with EBITDA at 10% a year."""
    a = with_case(simple(), revenue_growth=(0.10,) * 7, capex_pct=0.2)
    r = run_model(a).returns
    assert r.exit_ebitda == pytest.approx(100 * 1.1 ** 5)
    assert r.moic == pytest.approx(1.1 ** 5)
    assert r.irr == pytest.approx(0.10, abs=1e-12)


def test_option_pool_and_rollover_hand_case():
    a = simple(option_pool_pct=0.10, rollover_equity=100.0)
    r = run_model(a).returns
    assert r.sponsor_equity == pytest.approx(700.0)
    assert r.option_pool_value == pytest.approx(50.0)  # 10% of the 500 gain
    assert r.investor_equity == pytest.approx(1250.0)
    assert r.sponsor_proceeds == pytest.approx(700 / 800 * 1250)
    assert r.moic == pytest.approx(1250 / 800)
    assert r.rollover_proceeds == pytest.approx(100 / 800 * 1250)


def test_no_option_value_when_equity_falls():
    a = with_case(simple(option_pool_pct=0.2), exit_multiple=2.0)
    r = run_model(a).returns
    assert r.option_pool_value == 0.0
    assert r.moic < 1.0


def test_exit_multiple_override_beats_case():
    r = run_model(simple(), exit_multiple=10.0).returns
    assert r.exit_ev == pytest.approx(1000.0)


# ------------------------------------------------------------------ sources and uses
@pytest.mark.parametrize("key", SCENARIO_KEYS)
def test_sources_equal_uses(key):
    su = run_model(Assumptions(), key).su
    assert su["total_sources"] == pytest.approx(su["total_uses"])
    assert su["total_uses"] == pytest.approx(
        su["enterprise_value"] + su["transaction_fees"] + su["financing_fees"] + su["min_cash"])


def test_default_sources_and_uses_values():
    su = run_model(Assumptions()).su
    assert su["enterprise_value"] == pytest.approx(900.0)
    assert su["transaction_fees"] == pytest.approx(18.0)
    assert su["financing_fees"] == pytest.approx(0.025 * (60 + 400 + 150))
    assert su["sponsor_equity"] == pytest.approx(953.25 - 550 - 25)


def test_revolver_draw_at_close_is_a_source():
    su = run_model(Assumptions(revolver_draw_at_close=12.0)).su
    assert su["revolver"] == 12.0
    assert su["sponsor_equity"] == pytest.approx(953.25 - 12 - 550 - 25)


# ---------------------------------------------------------------- sweep and amortization
def _sweep_case(**over):
    return simple(**{"term_loan_x": 2.0, "sweep_pct": 1.0, **over})


def test_full_cash_sweep_hand_numbers():
    rows = run_model(_sweep_case()).rows
    assert rows["tl_close"][1] == pytest.approx(120.0)  # 100 - 20 interest = 80 swept
    assert rows["tl_close"][2] == pytest.approx(32.0)   # interest 12, sweep 88
    assert rows["tl_close"][3] == pytest.approx(0.0)    # interest 3.2, only 32 left to sweep
    assert rows["cash_close"][3] == pytest.approx(96.8 - 32.0)


def test_partial_cash_sweep_hand_numbers():
    rows = run_model(_sweep_case(sweep_pct=0.5)).rows
    assert rows["sweep"][1] == pytest.approx(40.0)
    assert rows["cash_close"][1] == pytest.approx(40.0)
    assert rows["tl_close"][2] == pytest.approx(160 - 42.0)   # sweep is 50% of the year's excess cash flow
    assert rows["cash_close"][2] == pytest.approx(40 + 84 - 42)


def test_zero_sweep_only_builds_cash():
    rows = run_model(_sweep_case(sweep_pct=0.0)).rows
    assert all(v == 200.0 for v in rows["tl_close"])
    assert rows["cash_close"][1] == pytest.approx(80.0)


def test_mandatory_amortization_is_pct_of_original():
    rows = run_model(_sweep_case(sweep_pct=0.0, mandatory_amort=(0.10,) * 7)).rows
    assert rows["mand"][1] == pytest.approx(20.0)
    assert rows["tl_close"][2] == pytest.approx(160.0)
    assert rows["cash_close"][1] == pytest.approx(100 - 20 - 20)


def test_mandatory_amortization_cannot_exceed_balance():
    rows = run_model(_sweep_case(sweep_pct=1.0, mandatory_amort=(0.60,) * 7)).rows
    assert rows["mand"][1] == pytest.approx(120.0)
    assert rows["mand"][2] == pytest.approx(80.0)   # only 80 left after year 1
    assert rows["tl_close"][2] == 0.0
    assert min(rows["tl_close"]) >= 0.0


def test_fully_repaid_debt_gives_zero_interest_and_na_coverage():
    res = run_model(_sweep_case())
    assert res.rows["tl_close"][3] == 0.0
    assert res.rows["gross_cash_interest"][4] == 0.0
    assert math.isnan(res.credit["interest_coverage"][4])
    assert math.isnan(res.credit["fccr"][4])
    assert res.credit["total_leverage"][4] == 0.0
    assert res.returns.exit_net_debt == pytest.approx(-res.rows["cash_close"][5])


def test_sweep_never_leaves_cash_below_minimum():
    res = run_model(Assumptions(), "base")
    assert min(res.rows["cash_close"]) >= Assumptions().min_cash - 1e-9


# ----------------------------------------------------------------------------- revolver
def _revolver_case(**over):
    a = with_case(simple(min_cash=20.0, revolver_commitment=100.0, sweep_pct=1.0), capex_pct=0.3)
    return replace(a, **over)


def test_revolver_funds_cash_shortfalls():
    rows = run_model(_revolver_case()).rows
    assert rows["rev_draw"][1] == pytest.approx(50.0)   # FCF is -50 (EBITDA 100, capex 150)
    assert rows["rev_close"][1] == pytest.approx(50.0)
    assert rows["cash_close"][1] == pytest.approx(20.0)


def test_revolver_is_capped_at_commitment_and_cash_breaches():
    rows = run_model(_revolver_case()).rows
    assert rows["rev_close"][2] == pytest.approx(100.0)
    assert rows["rev_draw"][3] == 0.0
    assert rows["cash_close"][3] < 20.0   # the checks sheet flags this


def test_revolver_repaid_before_sweep():
    a = with_case(simple(min_cash=0.0, revolver_commitment=100.0, revolver_draw_at_close=60.0,
                         term_loan_x=1.0, sweep_pct=1.0), capex_pct=0.0)
    rows = run_model(a).rows
    # year 1: interest on revolver 6 and term loan 10, FCF after interest 84 repays the revolver (60) first
    assert rows["rev_repay"][1] == pytest.approx(60.0)
    assert rows["rev_close"][1] == 0.0
    assert rows["sweep"][1] == pytest.approx(24.0)


def test_commitment_fee_on_undrawn_only():
    a = simple(revolver_commitment=100.0, commitment_fee=0.005)
    assert run_model(a).rows["int_commit"][1] == pytest.approx(0.5)


# ---------------------------------------------------------------------------------- PIK
def test_pik_accrues_to_principal_with_no_cash_interest():
    a = simple(sub_notes_x=1.0, notes_pik=True, notes_pik_rate=0.12)
    res = run_model(a)
    for t in range(8):
        assert res.rows["notes_close"][t] == pytest.approx(100 * 1.12 ** t)
    assert res.rows["gross_cash_interest"][1] == 0.0
    assert res.rows["int_notes_pik"][1] == pytest.approx(12.0)
    assert res.rows["cash_close"][1] == pytest.approx(100.0)   # no cash coupon paid


def test_cash_pay_notes_do_not_accrue():
    a = simple(sub_notes_x=1.0, notes_pik=False, notes_cash_rate=0.11)
    res = run_model(a)
    assert res.rows["notes_close"][5] == pytest.approx(100.0)
    assert res.rows["gross_cash_interest"][1] == pytest.approx(11.0)


def test_pik_interest_is_still_tax_deductible():
    cash = run_model(simple(sub_notes_x=1.0, notes_cash_rate=0.12, tax_rate=0.25))
    pik = run_model(simple(sub_notes_x=1.0, notes_pik=True, notes_pik_rate=0.12, tax_rate=0.25))
    assert pik.rows["tax"][1] == pytest.approx(cash.rows["tax"][1])


# ------------------------------------------------------------------------ taxes and fees
def test_taxes_floor_at_zero():
    a = simple(tax_rate=0.25, da_pct=0.5)   # D&A 250 on revenue 500 makes EBT negative
    rows = run_model(a).rows
    assert rows["ebt"][1] < 0
    assert rows["tax"][1] == 0.0


def test_tax_after_interest_hand_number():
    a = simple(term_loan_x=2.0, tax_rate=0.25)   # EBIT 100, interest 20, EBT 80, tax 20
    assert run_model(a).rows["tax"][1] == pytest.approx(20.0)


def test_financing_fee_amortization_is_deductible_and_stops():
    a = simple(term_loan_x=2.0, financing_fee_pct=0.025, fee_amort_years=2, tax_rate=0.25)
    rows = run_model(a).rows
    assert rows["fee_amort"][1] == pytest.approx(2.5)   # 5.0 of fees over two years
    assert rows["fee_amort"][3] == 0.0


# -------------------------------------------------------------------- credit statistics
def test_credit_statistics_hand_numbers():
    res = run_model(simple(term_loan_x=2.0, sub_notes_x=1.0, tax_rate=0.25, notes_cash_rate=0.10,
                           mandatory_amort=(0.05,) * 7))
    c = res.credit
    assert c["total_leverage"][0] == pytest.approx(3.0)
    assert c["senior_leverage"][0] == pytest.approx(2.0)
    gross = 20.0 + 10.0
    assert c["interest_coverage"][1] == pytest.approx(100 / gross)
    tax = 0.25 * (100 - gross)
    assert c["fccr"][1] == pytest.approx((100 - 0 - tax) / (gross + 10.0))


# ------------------------------------------------------------------------ average interest
def test_average_balance_interest_closed_form():
    """One year, TL 200 at 10%, 100% sweep: c = 100 + 0.05 * (200 + c)  =>  c = 110 / 0.95."""
    a = simple(term_loan_x=2.0, sweep_pct=1.0, interest_method="average")
    assert run_model(a).rows["tl_close"][1] == pytest.approx(110 / 0.95, abs=1e-9)


def test_average_method_converges_and_differs_from_opening():
    base = Assumptions()
    opening = run_model(base)
    average = run_model(base, interest_method="average")
    assert average.iterations > opening.iterations
    assert average.returns.irr > opening.returns.irr   # deleveraging lowers average interest
    assert abs(average.returns.irr - opening.returns.irr) < 0.01


def test_average_method_with_pik_is_a_fixed_point():
    res = run_model(simple(sub_notes_x=1.0, notes_pik=True, interest_method="average"))
    o, c = res.rows["notes_open"][1], res.rows["notes_close"][1]
    assert c - o == pytest.approx(0.12 * (o + c) / 2, abs=1e-9)


# ------------------------------------------------------------------------------- bridge
@pytest.mark.parametrize("key", SCENARIO_KEYS)
def test_bridge_reconciles_exactly(key):
    res = run_model(Assumptions(), key)
    assert abs(bridge_residual(res)) < 1e-9
    b = res.returns.bridge
    assert b["entry_equity"] + b["total_gain"] == pytest.approx(b["exit_equity"], abs=1e-9)


@pytest.mark.parametrize("exit_year", [3, 4, 5, 6, 7])
@pytest.mark.parametrize("pik", [False, True])
def test_bridge_reconciles_across_exit_years_and_pik(exit_year, pik):
    res = run_model(Assumptions(notes_pik=pik, exit_year=exit_year, revolver_draw_at_close=10.0), "downside")
    assert abs(bridge_residual(res)) < 1e-9
    r = res.returns
    assert r.sponsor_bridge["sponsor_gain"] == pytest.approx(r.sponsor_proceeds - r.sponsor_equity, abs=1e-9)


def test_bridge_hand_numbers():
    """Growth 10%, capex = EBITDA, exit at 10x versus 8x entry."""
    a = with_case(simple(), revenue_growth=(0.10,) * 7, capex_pct=0.2, exit_multiple=10.0)
    b = run_model(a).returns.bridge
    e5 = 100 * 1.1 ** 5
    assert b["ebitda_growth"] == pytest.approx(8 * (e5 - 100))
    assert b["multiple_expansion"] == pytest.approx(2 * e5)
    assert b["net_debt_paydown"] == pytest.approx(0.0)
    assert b["transaction_costs"] == 0.0
    assert b["exit_equity"] == pytest.approx(10 * e5)


def test_bridge_fees_are_a_negative_component():
    b = run_model(Assumptions()).returns.bridge
    assert b["transaction_costs"] == pytest.approx(-(18.0 + 15.25))


def test_net_debt_roll_forward_identity():
    res = run_model(Assumptions(notes_pik=True), "downside")
    rows = res.rows
    for t in range(1, 8):
        change = rows["net_debt"][t] - rows["net_debt"][t - 1]
        assert change == pytest.approx(rows["int_notes_pik"][t] - rows["fcf_post_int"][t], abs=1e-9)


# ---------------------------------------------------------------------------- scenarios
def test_scenario_ordering():
    irrs = {k: run_model(Assumptions(), k).returns.irr for k in SCENARIO_KEYS}
    assert irrs["management"] > irrs["base"] > irrs["downside"]


def test_downside_draws_the_revolver():
    assert max(run_model(Assumptions(), "downside").rows["rev_draw"]) > 0


def test_resolve_scenario_accepts_display_names():
    assert resolve_scenario("Management case") == "management"
    assert resolve_scenario(" Base ") == "base"
    with pytest.raises(ValueError):
        resolve_scenario("upside")


def test_validation_errors():
    with pytest.raises(ValueError):
        run_model(Assumptions(), exit_year=9)
    with pytest.raises(ValueError):
        run_model(Assumptions(), interest_method="median")
    with pytest.raises(ValueError):
        run_model(Assumptions(revolver_draw_at_close=80.0))


def test_shorter_horizon_runs():
    res = run_model(Assumptions(), horizon=5)
    assert len(res.rows["ebitda"]) == 6
    assert res.returns.exit_year == 5
