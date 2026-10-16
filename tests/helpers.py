"""Stripped-down assumption sets whose answers can be worked out by hand."""

from __future__ import annotations

from dataclasses import replace

from lbo_lab.assumptions import Assumptions, Case


def simple(**over) -> Assumptions:
    """Buy at 8x, flat 100 EBITDA on 500 revenue, no fees, no debt, no tax, no capex, no NWC."""
    case = Case("Base", (0.0,) * 7, (0.2,) * 7, 0.0, 8.0)
    base = Assumptions(
        ltm_revenue=500.0, ltm_ebitda=100.0, entry_multiple=8.0, transaction_fee_pct=0.0,
        financing_fee_pct=0.0, min_cash=0.0, rollover_equity=0.0, option_pool_pct=0.0,
        revolver_commitment=0.0, term_loan_x=0.0, sub_notes_x=0.0, tax_rate=0.0, da_pct=0.0,
        nwc_pct=0.0, interest_income_rate=0.0, commitment_fee=0.0, sweep_pct=0.0,
        mandatory_amort=(0.0,) * 7, exit_year=5, cases={"base": case},
        base_rate=0.10, term_spread=0.0, revolver_spread=0.0,
    )
    return replace(base, **over)


def with_case(a: Assumptions, **case_over) -> Assumptions:
    return replace(a, cases={"base": replace(a.cases["base"], **case_over)})
