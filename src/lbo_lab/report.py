"""Plain-text reports for the CLI."""

from __future__ import annotations

import math

from .engine import ModelResult
from .sensitivity import Grid

RULE = "-" * 86


def _f(x: float, nd: int = 1) -> str:
    return "n.a." if isinstance(x, float) and math.isnan(x) else f"{x:,.{nd}f}"


def _x(x: float) -> str:
    return "n.a." if math.isnan(x) else f"{x:.2f}x"


def _pct(x: float) -> str:
    return f"{100 * x:.1f}%"


def format_run(res: ModelResult) -> str:
    a, su, r, n = res.assumptions, res.su, res.returns, res.n
    ebitda0 = a.ltm_ebitda
    out = [
        f"{a.company} | {a.case.name} case | exit year {r.exit_year} at {r.exit_multiple:.1f}x | "
        f"interest on {a.interest_method} balances",
        "USD millions. Fictional company, illustrative figures.",
        "",
        "SOURCES AND USES",
        RULE,
        f"{'Sources':<28}{'$mm':>9}{'%':>7}{'x EBITDA':>10}   {'Uses':<28}{'$mm':>9}{'%':>7}",
    ]
    src = [("Revolver (drawn)", su["revolver"]), ("Senior term loan", su["term_loan"]),
           ("Subordinated notes", su["sub_notes"]), ("Management rollover", su["rollover"]),
           ("Sponsor equity", su["sponsor_equity"]), ("Total sources", su["total_sources"])]
    use = [("Enterprise value", su["enterprise_value"]), ("Transaction fees", su["transaction_fees"]),
           ("Financing fees", su["financing_fees"]), ("Minimum cash", su["min_cash"]),
           ("", None), ("Total uses", su["total_uses"])]
    for (sn, sv), (un, uv) in zip(src, use, strict=True):
        left = f"{sn:<28}{_f(sv):>9}{_pct(sv / su['total_sources']):>7}{sv / ebitda0:>9.2f}x"
        right = f"{un:<28}{_f(uv):>9}{_pct(uv / su['total_uses']):>7}" if uv is not None else ""
        out.append(f"{left}   {right}")
    out += ["", "RETURNS", RULE,
            f"Exit EBITDA {_f(r.exit_ebitda)} x {r.exit_multiple:.1f}x = exit EV {_f(r.exit_ev)}; "
            f"net debt {_f(r.exit_net_debt)}; exit equity {_f(r.exit_equity)}",
            f"Option pool {_f(r.option_pool_value)}; sponsor owns {_pct(r.sponsor_ownership)} of investor equity; "
            f"proceeds {_f(r.sponsor_proceeds)} on {_f(r.sponsor_equity)} invested",
            f"Sponsor MOIC {r.moic:.2f}x    Sponsor IRR {_pct(r.irr)}", "",
            "VALUE CREATION BRIDGE (100% of equity)", RULE]
    b = r.bridge
    gain = b["total_gain"]
    out.append(f"{'Entry equity':<34}{_f(b['entry_equity']):>10}")
    for key, text in (("ebitda_growth", "EBITDA growth (at entry multiple)"),
                      ("multiple_expansion", "Multiple expansion (on exit EBITDA)"),
                      ("net_debt_paydown", "Net debt paydown"),
                      ("transaction_costs", "Transaction and financing costs")):
        share = b[key] / gain if gain else 0.0
        out.append(f"{text:<34}{_f(b[key]):>10}{_pct(share):>9} of gain")
    out.append(f"{'Exit equity (entry + gain)':<34}{_f(b['exit_equity']):>10}   (gain {_f(gain)})")
    out += ["", "OPERATING AND CREDIT PROFILE", RULE]
    hdr = f"{'':<30}" + "".join(f"{'Close' if t == 0 else f'FY{a.base_year + t}E':>9}" for t in range(n + 1))
    out.append(hdr)
    rows = res.rows
    table = [("Revenue", rows["revenue"], _f), ("EBITDA", rows["ebitda"], _f),
             ("FCF before debt service", rows["cfads"], _f), ("FCF after debt service", rows["fcf_post_mand"], _f),
             ("Total debt", rows["total_debt"], _f), ("Cash", rows["cash_close"], _f),
             ("Total debt / EBITDA", res.credit["total_leverage"], _x),
             ("Senior debt / EBITDA", res.credit["senior_leverage"], _x),
             ("Interest coverage", res.credit["interest_coverage"], _x),
             ("FCCR", res.credit["fccr"], _x)]
    for name, vals, fn in table:
        cells = ["" if (t == 0 and name in ("FCF before debt service", "FCF after debt service",
                                              "Interest coverage", "FCCR")) else fn(vals[t])
                 for t in range(n + 1)]
        out.append(f"{name:<30}" + "".join(f"{c:>9}" for c in cells))
    return "\n".join(out)


def format_grid(grid: Grid, metric: str = "irr") -> str:
    data = grid.irr if metric == "irr" else grid.moic
    cell = (lambda v: f"{100 * v:.1f}%") if metric == "irr" else (lambda v: f"{v:.2f}x")
    col_fmt = (lambda v: f"{v:.0f}") if grid.col_label == "Exit year" else (lambda v: f"{v:.1f}x")
    row_fmt = (lambda v: f"{v:.1f}x")
    head = (f"{grid.name} | sponsor {'IRR' if metric == 'irr' else 'MOIC'} | "
            f"rows: {grid.row_label}, columns: {grid.col_label}")
    lines = [head, f"{'':>10}" + "".join(f"{col_fmt(c):>9}" for c in grid.col_values)]
    for rv, row in zip(grid.row_values, data, strict=True):
        lines.append(f"{row_fmt(rv):>10}" + "".join(f"{cell(v):>9}" for v in row))
    return "\n".join(lines)
