"""Leveraged buyout engine: sources and uses, projections, debt waterfall, returns, bridge.

Timeline convention: index 0 is the closing balance sheet / LTM period, indices 1..n are
projection years. Flow items carry 0.0 at index 0. Interest is charged on opening
balances by default (this mirrors the Excel workbook, which has no circular references).
With ``interest_method="average"`` interest is charged on average balances and each
year is solved to a fixed point by iteration.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field, replace

from .assumptions import Assumptions

NAN = float("nan")
MAX_ITER = 500
TOL = 1e-13


# ----------------------------------------------------------------------------- IRR


def irr(cashflows: list[float], lo: float = -0.9999, hi: float = 10.0) -> float:
    """Internal rate of return by bisection on annual cash flows (index = year)."""
    if not (any(c < 0 for c in cashflows) and any(c > 0 for c in cashflows)):
        return NAN

    def npv(rate: float) -> float:
        return sum(c / (1.0 + rate) ** t for t, c in enumerate(cashflows))

    f_lo, f_hi = npv(lo), npv(hi)
    if f_lo * f_hi > 0:
        return NAN
    for _ in range(300):
        mid = 0.5 * (lo + hi)
        f_mid = npv(mid)
        if f_lo * f_mid <= 0:
            hi = mid
        else:
            lo, f_lo = mid, f_mid
        if hi - lo < 1e-14:
            break
    return 0.5 * (lo + hi)


def irr_from_moic(moic: float, years: int) -> float:
    """Closed-form IRR for one outflow at t=0 and one inflow at t=years."""
    if moic <= 0:
        return -1.0
    return moic ** (1.0 / years) - 1.0


# --------------------------------------------------------------------- result types


@dataclass
class Returns:
    exit_year: int
    exit_multiple: float
    exit_ebitda: float
    exit_ev: float
    exit_debt: float
    exit_cash: float
    exit_net_debt: float
    exit_equity: float
    entry_equity: float
    sponsor_equity: float
    option_pool_value: float
    investor_equity: float
    sponsor_ownership: float
    sponsor_proceeds: float
    rollover_proceeds: float
    moic: float
    irr: float
    irr_from_flows: float
    sponsor_cash_flows: list[float]
    bridge: dict[str, float]
    sponsor_bridge: dict[str, float]


@dataclass
class ModelResult:
    assumptions: Assumptions
    su: dict[str, float]
    rows: dict[str, list[float]]
    credit: dict[str, list[float]]
    returns: Returns
    iterations: int = 1
    notes: list[str] = field(default_factory=list)

    @property
    def n(self) -> int:
        return self.assumptions.horizon


# ------------------------------------------------------------------ sources and uses


def sources_and_uses(a: Assumptions) -> dict[str, float]:
    ev = a.entry_multiple * a.ltm_ebitda
    term_loan = a.term_loan_x * a.ltm_ebitda
    notes = a.sub_notes_x * a.ltm_ebitda
    txn_fees = a.transaction_fee_pct * ev
    fin_fees = a.financing_fee_pct * (a.revolver_commitment + term_loan + notes)
    uses = ev + txn_fees + fin_fees + a.min_cash
    sponsor = uses - a.revolver_draw_at_close - term_loan - notes - a.rollover_equity
    sources = a.revolver_draw_at_close + term_loan + notes + a.rollover_equity + sponsor
    debt = a.revolver_draw_at_close + term_loan + notes
    return {
        "enterprise_value": ev,
        "transaction_fees": txn_fees,
        "financing_fees": fin_fees,
        "min_cash": a.min_cash,
        "total_uses": uses,
        "revolver": a.revolver_draw_at_close,
        "term_loan": term_loan,
        "sub_notes": notes,
        "rollover": a.rollover_equity,
        "sponsor_equity": sponsor,
        "total_sources": sources,
        "total_debt": debt,
        "total_equity": sponsor + a.rollover_equity,
    }


# ------------------------------------------------------------------- annual solver


def _year(a: Assumptions, op: dict[str, float], t: int, su: dict[str, float],
          opn: dict[str, float], basis: dict[str, float]) -> dict[str, float]:
    """Solve one projection year given opening balances and an interest basis."""
    rev_rate = a.base_rate + a.revolver_spread
    tl_rate = a.base_rate + a.term_spread
    notes_rate = a.notes_pik_rate if a.notes_pik else a.notes_cash_rate

    int_rev = rev_rate * basis["rev"]
    int_commit = a.commitment_fee * (a.revolver_commitment - basis["rev"])
    int_tl = tl_rate * basis["tl"]
    int_notes = notes_rate * basis["notes"]
    int_notes_cash = 0.0 if a.notes_pik else int_notes
    int_notes_pik = int_notes if a.notes_pik else 0.0
    int_income = a.interest_income_rate * basis["cash"]
    gross_cash_interest = int_rev + int_commit + int_tl + int_notes_cash
    net_cash_interest = gross_cash_interest - int_income
    total_interest = int_rev + int_commit + int_tl + int_notes

    fee_amort = su["financing_fees"] / a.fee_amort_years if t <= a.fee_amort_years else 0.0
    ebt = op["ebit"] - fee_amort - total_interest + int_income
    tax = max(0.0, ebt) * a.tax_rate
    cfads = op["ebitda"] - tax - op["capex"] - op["dnwc"]
    fcf_post_int = cfads - net_cash_interest

    mand = min(opn["tl"], a.mandatory_amort[t - 1] * su["term_loan"])
    fcf_post_mand = fcf_post_int - mand
    cash_pre = opn["cash"] + fcf_post_mand
    avail = cash_pre - a.min_cash
    rev_draw = min(max(-avail, 0.0), a.revolver_commitment - opn["rev"])
    rev_repay = min(opn["rev"], max(avail, 0.0))
    avail2 = avail + rev_draw - rev_repay
    ecf = max(0.0, fcf_post_mand - rev_repay)
    sweep = max(0.0, min(opn["tl"] - mand, a.sweep_pct * ecf, avail2))

    return {
        "int_rev": int_rev, "int_commit": int_commit, "int_tl": int_tl,
        "int_notes": int_notes, "int_notes_cash": int_notes_cash,
        "int_notes_pik": int_notes_pik, "int_income": int_income,
        "gross_cash_interest": gross_cash_interest, "net_cash_interest": net_cash_interest,
        "total_interest": total_interest, "fee_amort": fee_amort, "ebt": ebt, "tax": tax,
        "net_income": ebt - tax, "cfads": cfads, "fcf_post_int": fcf_post_int,
        "mand": mand, "fcf_post_mand": fcf_post_mand, "cash_pre": cash_pre,
        "avail": avail, "avail2": avail2, "ecf": ecf,
        "rev_draw": rev_draw, "rev_repay": rev_repay, "sweep": sweep,
        "rev_close": opn["rev"] + rev_draw - rev_repay,
        "tl_close": opn["tl"] - mand - sweep,
        "notes_close": opn["notes"] + int_notes_pik,
        "cash_close": cash_pre + rev_draw - rev_repay - sweep,
    }


ROW_KEYS = (
    "revenue growth ebitda margin da ebit capex dnwc fee_amort int_rev int_commit int_tl "
    "int_notes int_notes_cash int_notes_pik int_income gross_cash_interest net_cash_interest "
    "total_interest ebt tax net_income cfads fcf_post_int mand fcf_post_mand cash_pre avail "
    "avail2 ecf rev_draw rev_repay sweep rev_open rev_close tl_open tl_close notes_open "
    "notes_close cash_open cash_close total_debt senior_debt net_debt"
).split()


def run_model(a: Assumptions, scenario: str | None = None, **overrides: object) -> ModelResult:
    """Run the full LBO. Keyword overrides replace Assumptions fields for this run only."""
    if scenario is not None:
        a = a.with_scenario(scenario)
    if overrides:
        a = replace(a, **overrides)
    a.validate()
    case, n = a.case, a.horizon
    su = sources_and_uses(a)
    rows: dict[str, list[float]] = {k: [0.0] * (n + 1) for k in ROW_KEYS}

    rows["revenue"][0] = a.ltm_revenue
    rows["ebitda"][0] = a.ltm_ebitda
    rows["margin"][0] = a.ltm_ebitda / a.ltm_revenue
    rows["rev_close"][0] = su["revolver"]
    rows["tl_close"][0] = su["term_loan"]
    rows["notes_close"][0] = su["sub_notes"]
    rows["cash_close"][0] = su["min_cash"]

    total_iter = 0
    for t in range(1, n + 1):
        g, m = case.revenue_growth[t - 1], case.ebitda_margin[t - 1]
        rev = rows["revenue"][t - 1] * (1.0 + g)
        op = {
            "ebitda": rev * m,
            "capex": case.capex_pct * rev,
            "dnwc": a.nwc_pct * (rev - rows["revenue"][t - 1]),
        }
        op["ebit"] = op["ebitda"] - a.da_pct * rev
        rows["growth"][t], rows["revenue"][t], rows["margin"][t] = g, rev, m
        rows["ebitda"][t], rows["capex"][t] = op["ebitda"], op["capex"]
        rows["dnwc"][t], rows["ebit"][t], rows["da"][t] = op["dnwc"], op["ebit"], a.da_pct * rev

        opn = {"rev": rows["rev_close"][t - 1], "tl": rows["tl_close"][t - 1],
               "notes": rows["notes_close"][t - 1], "cash": rows["cash_close"][t - 1]}
        basis = dict(opn)
        for _ in range(MAX_ITER):
            y = _year(a, op, t, su, opn, basis)
            total_iter += 1
            if a.interest_method == "opening":
                break
            new_basis = {k: 0.5 * (opn[k] + y[f"{k}_close"]) for k in opn}
            err = max(abs(new_basis[k] - basis[k]) for k in opn)
            basis = new_basis
            if err < TOL:
                break
        else:
            raise RuntimeError(f"average-balance interest did not converge in year {t}")

        for k, v in y.items():
            if k in rows:
                rows[k][t] = v
        rows["rev_open"][t], rows["tl_open"][t] = opn["rev"], opn["tl"]
        rows["notes_open"][t], rows["cash_open"][t] = opn["notes"], opn["cash"]

    for t in range(n + 1):
        rows["total_debt"][t] = rows["rev_close"][t] + rows["tl_close"][t] + rows["notes_close"][t]
        rows["senior_debt"][t] = rows["rev_close"][t] + rows["tl_close"][t]
        rows["net_debt"][t] = rows["total_debt"][t] - rows["cash_close"][t]

    credit = _credit_stats(rows, n)
    ret = _returns(a, su, rows)
    return ModelResult(a, su, rows, credit, ret, iterations=total_iter)


# ----------------------------------------------------------------- credit statistics


def _div(num: float, den: float) -> float:
    return num / den if den > 1e-9 else NAN


def _credit_stats(rows: dict[str, list[float]], n: int) -> dict[str, list[float]]:
    keys = ["total_leverage", "senior_leverage", "net_leverage",
            "interest_coverage", "capex_coverage", "fccr"]
    out: dict[str, list[float]] = {k: [NAN] * (n + 1) for k in keys}
    for t in range(n + 1):
        e = rows["ebitda"][t]
        out["total_leverage"][t] = rows["total_debt"][t] / e
        out["senior_leverage"][t] = rows["senior_debt"][t] / e
        out["net_leverage"][t] = rows["net_debt"][t] / e
        if t == 0:
            continue
        gci = rows["gross_cash_interest"][t]
        out["interest_coverage"][t] = _div(e, gci)
        out["capex_coverage"][t] = _div(e - rows["capex"][t], gci)
        fixed = gci + rows["mand"][t]
        out["fccr"][t] = _div(e - rows["capex"][t] - rows["tax"][t], fixed)
    return out


# ------------------------------------------------------------------------- returns


def _returns(a: Assumptions, su: dict[str, float], rows: dict[str, list[float]]) -> Returns:
    n_exit = a.exit_year
    mult = a.effective_exit_multiple
    ebitda_x = rows["ebitda"][n_exit]
    ev_x = mult * ebitda_x
    debt_x = rows["total_debt"][n_exit]
    cash_x = rows["cash_close"][n_exit]
    nd_x = debt_x - cash_x
    equity_x = ev_x - nd_x

    e0 = su["total_equity"]
    sponsor_eq = su["sponsor_equity"]
    pool = a.option_pool_pct * max(0.0, equity_x - e0)
    investors = equity_x - pool
    own = sponsor_eq / e0
    proceeds = own * investors
    moic = proceeds / sponsor_eq
    flows = [0.0] * (a.horizon + 1)
    flows[0] = -sponsor_eq
    flows[n_exit] += proceeds

    nd0 = su["total_debt"] - su["min_cash"]
    fees = su["transaction_fees"] + su["financing_fees"]
    bridge = {
        "entry_equity": e0,
        "ebitda_growth": a.entry_multiple * (ebitda_x - a.ltm_ebitda),
        "multiple_expansion": (mult - a.entry_multiple) * ebitda_x,
        "net_debt_paydown": nd0 - nd_x,
        "transaction_costs": -fees,
    }
    bridge["total_gain"] = sum(v for k, v in bridge.items() if k != "entry_equity")
    bridge["exit_equity"] = equity_x
    sponsor_bridge = {
        "option_pool_dilution": -pool,
        "investor_gain": bridge["total_gain"] - pool,
        "sponsor_ownership": own,
        "sponsor_gain": own * (bridge["total_gain"] - pool),
    }
    return Returns(
        exit_year=n_exit, exit_multiple=mult, exit_ebitda=ebitda_x, exit_ev=ev_x,
        exit_debt=debt_x, exit_cash=cash_x, exit_net_debt=nd_x, exit_equity=equity_x,
        entry_equity=e0, sponsor_equity=sponsor_eq, option_pool_value=pool,
        investor_equity=investors, sponsor_ownership=own, sponsor_proceeds=proceeds,
        rollover_proceeds=investors - proceeds, moic=moic,
        irr=irr_from_moic(moic, n_exit), irr_from_flows=irr(flows),
        sponsor_cash_flows=flows, bridge=bridge, sponsor_bridge=sponsor_bridge,
    )


def bridge_residual(res: ModelResult) -> float:
    """Total gain implied by the bridge parts minus (exit equity - entry equity)."""
    r = res.returns
    parts = sum(v for k, v in r.bridge.items() if k not in ("entry_equity", "total_gain", "exit_equity"))
    return parts - (r.exit_equity - r.entry_equity)


def is_nan(x: float) -> bool:
    return isinstance(x, float) and math.isnan(x)
