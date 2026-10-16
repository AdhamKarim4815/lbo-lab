"""Command line interface: ``lbo run | build | sens | charts | scenarios``."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import replace
from pathlib import Path

from . import __version__, charts, excel, recalc, report, sensitivity
from .assumptions import SCENARIO_KEYS, Assumptions, resolve_scenario
from .engine import run_model


def _common(p: argparse.ArgumentParser) -> None:
    p.add_argument("--scenario", default="base", help="base, management or downside (default: base)")
    p.add_argument("--exit-year", type=int, help="exit year, 1 to 7 (default 5)")
    p.add_argument("--exit-multiple", type=float, help="override the case exit EV/EBITDA")
    p.add_argument("--pik", action="store_true", help="subordinated notes pay in kind")


def _assumptions(args: argparse.Namespace) -> Assumptions:
    a = Assumptions().with_scenario(resolve_scenario(args.scenario))
    over: dict[str, object] = {}
    if args.exit_year is not None:
        over["exit_year"] = args.exit_year
    if args.exit_multiple is not None:
        over["exit_multiple"] = args.exit_multiple
    if args.pik:
        over["notes_pik"] = True
    if getattr(args, "interest_method", None):
        over["interest_method"] = args.interest_method
    return replace(a, **over)


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="lbo", description="Halcyon Packaging LBO model (fictional company).")
    p.add_argument("--version", action="version", version=f"lbo-lab {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    run = sub.add_parser("run", help="print sources and uses, returns, bridge and credit statistics")
    _common(run)
    run.add_argument("--interest-method", choices=["opening", "average"],
                     help="interest on opening balances (default) or iterated average balances")
    run.add_argument("--json", action="store_true", help="emit key results as JSON")

    build = sub.add_parser("build", help="write the Excel workbook with live formulas")
    _common(build)
    build.add_argument("--out", default="lbo_model.xlsx", help="output path (default lbo_model.xlsx)")
    build.add_argument("--no-recalc", action="store_true", help="skip the LibreOffice pass that caches values")

    sens = sub.add_parser("sens", help="print sensitivity tables")
    _common(sens)
    sens.add_argument("--grid", choices=["entry", "leverage", "exit", "all"], default="all")

    ch = sub.add_parser("charts", help="write PNG charts")
    _common(ch)
    ch.add_argument("--out-dir", default="docs", help="output directory (default docs)")

    sub.add_parser("scenarios", help="list the available scenarios")
    return p


def _json(res) -> str:
    r = res.returns
    return json.dumps({
        "scenario": res.assumptions.case.name, "exit_year": r.exit_year, "sponsor_irr": r.irr,
        "sponsor_moic": r.moic, "exit_equity": r.exit_equity, "sponsor_equity": r.sponsor_equity,
        "sources_and_uses": res.su, "bridge": r.bridge,
        "total_leverage": res.credit["total_leverage"], "ending_debt": res.rows["total_debt"],
    }, indent=2)


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "scenarios":
            a = Assumptions()
            for k in SCENARIO_KEYS:
                c = a.cases[k]
                print(f"{k:<11} exit {c.exit_multiple:.1f}x  capex {100 * c.capex_pct:.1f}%  "
                      f"year-1 growth {100 * c.revenue_growth[0]:.1f}%  margin {100 * c.ebitda_margin[0]:.1f}%")
            return 0
        a = _assumptions(args)
        a.validate()
        if args.command == "run":
            res = run_model(a)
            print(_json(res) if args.json else report.format_run(res))
        elif args.command == "build":
            out = Path(args.out)
            out.parent.mkdir(parents=True, exist_ok=True)
            excel.build_model_file(out, a, recalculate=not args.no_recalc)
            cached = not args.no_recalc and recalc.find_soffice() is not None
            print(f"Wrote {out} ({a.case.name} case; cached values: {'yes' if cached else 'no'})")
        elif args.command == "sens":
            grids = {"entry": sensitivity.entry_vs_exit(a), "leverage": sensitivity.leverage_vs_exit_year(a),
                     "exit": sensitivity.exit_year_vs_exit_multiple(a)}
            for key, grid in grids.items():
                if args.grid in (key, "all"):
                    print(report.format_grid(grid, "irr"), "\n")
                    print(report.format_grid(grid, "moic"), "\n")
        elif args.command == "charts":
            res = run_model(a)
            paths = charts.make_all(res, sensitivity.entry_vs_exit(a), Path(args.out_dir))
            print("\n".join(f"Wrote {p}" for p in paths))
    except (ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
