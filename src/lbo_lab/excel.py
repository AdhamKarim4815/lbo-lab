"""Banker-formatted Excel LBO workbook with live formulas.

Every number outside the Inputs sheet is a formula. Interest is charged on opening
balances, so the workbook has no circular references and recalculates in Excel and
LibreOffice alike. ``build_model_file`` writes the workbook, recalculates it headlessly
with LibreOffice and writes it a second time with the cached results, so viewers that do
not recalculate (previews, openpyxl) still show numbers.
"""

from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import xlsxwriter

from . import recalc
from .assumptions import HORIZON, SCENARIO_KEYS, Assumptions

NAVY, STEEL, GREY, LIGHT = "#1F3864", "#2E75B6", "#7F7F7F", "#EEF2F8"
BLUE, BLACK, GREEN, WHITE = "#0000FF", "#000000", "#008000", "#FFFFFF"
INPUT_BG, GOOD, BAD = "#FFF8DC", "#C6E0B4", "#F4B6B6"
YC = list("DEFGHIJK")  # column of year t, t = 0..7
SHEETS = ["Cover", "Summary", "Inputs", "Transaction", "Operating", "Debt", "Returns",
          "Sensitivity", "Checks"]
NUM = {
    "gen": "General",
    "num": '#,##0.0_);(#,##0.0);"-"_)',
    "num0": '#,##0_);(#,##0);"-"_)',
    "pct": '0.0%_);(0.0%);"-"_)',
    "pct2": '0.00%_);(0.00%);"-"_)',
    "mult": '0.0"x"_)',
    "mult2": '0.00"x"_)',
    "int": "0_)",
    "year": '"Year "0',
    "fy": '"FY"0"E"',
    "fy0": '"LTM FY"0',
    "sci": "0.0E+00",
}
LINK_RE = re.compile(r"^=(?:INDEX\()?[A-Za-z]+!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?"
                     r"(?:,\$?[A-Z]+\$?\d+)?\)?$")
CALC_SHEETS = ("Transaction", "Operating", "Debt", "Returns", "Sensitivity", "Summary", "Checks")


@dataclass
class Layout:
    """Where things live in the workbook; used by tests and the docs script."""

    rows: dict[str, dict[str, int]] = field(default_factory=dict)
    grids: dict[str, dict[str, object]] = field(default_factory=dict)
    header_rows: dict[str, list[int]] = field(default_factory=dict)

    def cell(self, sheet: str, key: str, col: str = "D") -> str:
        return f"{col}{self.rows[sheet][key]}"


def _rows(spec: list[str | None], start: int) -> dict[str, int]:
    out, r = {}, start
    for k in spec:
        if k is not None:
            out[k] = r
        r += 1
    return out


class Book:
    def __init__(self, path: Path, a: Assumptions, cache: dict | None = None) -> None:
        if a.horizon != HORIZON:
            raise ValueError(f"the Excel workbook is built for a {HORIZON}-year horizon")
        a.validate()
        self.a, self.cache = a, cache or {}
        self.wb = xlsxwriter.Workbook(str(path))
        self.wb.set_properties({"title": "Halcyon Packaging Co. LBO model",
                                "author": "Adham Karim",
                                "comments": "Fictional company. Illustrative model."})
        self.ws = {n: self.wb.add_worksheet(n) for n in SHEETS}
        self.r: dict[str, dict[str, int]] = {}
        self.layout = Layout()
        self._fmts: dict[tuple, object] = {}
        for ws in self.ws.values():
            ws.hide_gridlines(2)
            ws.set_column("A:A", 2)
            ws.set_column("B:B", 46)
            ws.set_column("C:C", 14)
            ws.set_column("D:K", 12.5)
            ws.set_landscape()
            ws.fit_to_pages(1, 0)

    # ----------------------------------------------------------------- primitives
    def fmt(self, num: str = "gen", color: str = BLACK, bold: bool = False, italic: bool = False,
            bg: str | None = None, top: bool = False, bottom: bool = False,
            align: str | None = None, size: int = 10, wrap: bool = False, indent: int = 0,
            box: bool = False):
        key = (num, color, bold, italic, bg, top, bottom, align, size, wrap, indent, box)
        if key not in self._fmts:
            d: dict[str, object] = {"font_name": "Calibri", "font_size": size,
                                    "font_color": color, "num_format": NUM.get(num, num)}
            if bold:
                d["bold"] = True
            if italic:
                d["italic"] = True
            if bg:
                d["bg_color"] = bg
            if top:
                d["top"] = 1
            if bottom:
                d["bottom"] = 1
            if align:
                d["align"] = align
            if wrap:
                d["text_wrap"] = True
                d["valign"] = "top"
            if indent:
                d["indent"] = indent
            if box:
                d["border"] = 1
                d["border_color"] = "#BFBFBF"
            self._fmts[key] = self.wb.add_format(d)
        return self._fmts[key]

    def sheet(self, name: str, spec: list[str | None], start: int) -> dict[str, int]:
        self.r[name] = _rows(spec, start)
        self.layout.rows[name] = self.r[name]
        return self.r[name]

    def A(self, sh: str, key: str, col: str = "D", absolute: bool = False, ext: bool = False) -> str:
        row = self.r[sh][key]
        cell = f"${col}${row}" if absolute else f"{col}{row}"
        return f"{sh}!{cell}" if ext else cell

    def X(self, sh: str, key: str, col: str = "D") -> str:
        """Absolute cross-sheet reference."""
        return self.A(sh, key, col, absolute=True, ext=True)

    def inp(self, sh: str, cell: str, value, num: str = "num", **kw) -> None:
        self.ws[sh].write(cell, value, self.fmt(num, BLUE, bg=INPUT_BG, box=True, **kw))

    def f(self, sh: str, cell: str, formula: str, num: str = "num", **kw) -> None:
        color = GREEN if LINK_RE.match(formula) else BLACK
        color = kw.pop("color", color)
        value = self.cache.get((sh, cell))
        fm = self.fmt(num, color, **kw)
        if value is None:
            self.ws[sh].write_formula(cell, formula, fm)
        else:
            self.ws[sh].write_formula(cell, formula, fm, value)

    def t(self, sh: str, cell: str, text: str, **kw) -> None:
        self.ws[sh].write_string(cell, text, self.fmt("gen", **kw))

    def band(self, sh: str, row: int, text: str, last: str = "K") -> None:
        fm = self.fmt(color=WHITE, bold=True, bg=NAVY)
        for col in range(1, ord(last) - 64):
            self.ws[sh].write_blank(row - 1, col, None, fm)
        self.ws[sh].write_string(f"B{row}", text, fm)

    def title(self, sh: str, text: str, sub: str) -> None:
        self.t(sh, "B1", text, bold=True, size=15, color=NAVY)
        self.t(sh, "B2", sub, italic=True, color=GREY)

    def lab(self, sh: str, key: str, text: str, unit: str = "", bold: bool = False,
            indent: int = 0, top: bool = False) -> None:
        row = self.r[sh][key]
        self.t(sh, f"B{row}", text, bold=bold, indent=indent, top=top)
        if unit or top:
            self.t(sh, f"C{row}", unit, color=GREY, italic=True, top=top)

    def ser(self, sh: str, key: str, fn, t0: int = 1, t1: int = 7, num: str = "num", **kw) -> None:
        row = self.r[sh][key]
        for t in range(t0, t1 + 1):
            self.f(sh, f"{YC[t]}{row}", fn(YC[t], YC[t - 1], t), num, **kw)

    def one(self, sh: str, key: str, formula: str, num: str = "num", col: str = "D", **kw) -> None:
        self.f(sh, f"{col}{self.r[sh][key]}", formula, num, **kw)

    def year_header(self, sh: str, yr: str, fy: str) -> None:
        ry, rf = self.r[sh][yr], self.r[sh][fy]
        self.layout.header_rows.setdefault(sh, []).append(ry)
        hdr = {"color": WHITE, "bold": True, "bg": NAVY, "align": "right"}
        self.t(sh, f"B{ry}", "Projection year", color=WHITE, bold=True, bg=NAVY)
        self.t(sh, f"C{ry}", "", color=WHITE, bg=NAVY)
        self.t(sh, f"B{rf}", "Fiscal year", color=WHITE, bold=True, bg=NAVY)
        self.t(sh, f"C{rf}", "", color=WHITE, bg=NAVY)
        for t in range(8):
            c = YC[t]
            self.f(sh, f"{c}{ry}", "=COLUMN()-COLUMN($D$1)", "year", **hdr)
            self.f(sh, f"{c}{rf}", f"=$D${self.r[sh]['base_year']}+{c}{ry}",
                   "fy0" if t == 0 else "fy", **hdr)

    def link_block(self, sh: str, items: list[tuple[str, str, str, str]],
                   bold: tuple[str, ...] = ()) -> None:
        """items: (key, label, formula, numfmt) written in column D."""
        for key, text, formula, num in items:
            self.lab(sh, key, text, bold=key in bold)
            self.one(sh, key, formula, num, bold=key in bold, bg=LIGHT if key in bold else None)

    # ---------------------------------------------------------------------- build
    def build(self) -> Layout:
        self._layouts()
        self.inputs()
        self.transaction()
        self.operating()
        self.debt()
        self.returns()
        self.sensitivity()
        self.checks()
        self.summary()
        self.cover()
        self.wb.close()
        return self.layout

    def _layouts(self) -> None:
        self.sheet("Inputs", [
            "h_scn", "scenario", "scn_idx", "scn1", "scn2", "scn3", None,
            "h_txn", "company", "base_year", "ltm_revenue", "ltm_ebitda", "entry_mult",
            "txn_fee_pct", "fin_fee_pct", "fee_years", "min_cash", "rollover", "pool_pct", None,
            "h_fin", "rev_commit", "rev_draw", "tl_x", "notes_x", "base_rate", "rev_spread",
            "tl_spread", "commit_fee", "notes_cash", "notes_pik_rate", "pik_toggle",
            "int_income", "sweep_pct", "mand_hdr", "mand_amort", None,
            "h_op", "da_pct", "nwc_pct", "tax_rate", None,
            "h_exit", "exit_year", "exit_mult", "sens_step", None,
            "h_cases", "case_hdr",
            "c0_name", "c0_growth", "c0_margin", "c0_capex", "c0_exit", None,
            "c1_name", "c1_growth", "c1_margin", "c1_capex", "c1_exit", None,
            "c2_name", "c2_growth", "c2_margin", "c2_capex", "c2_exit", None,
            "h_live", "live_growth", "live_margin", "live_capex", "live_exit"], 4)
        self.sheet("Transaction", [
            "h_pp", "ltm_ebitda", "entry_mult", "ev", "txn_fee_pct", "txn_fees", "fin_fee_pct",
            "rev_commit", "fin_base", "fin_fees", "min_cash", None,
            "h_s", "s_hdr", "s_rev", "s_tl", "s_notes", "s_roll", "s_sponsor", "s_total", None,
            "h_u", "u_hdr", "u_ev", "u_txn", "u_fin", "u_cash", "u_total", None, "su_check", None,
            "h_cap", "cap_hdr", "c_rev", "c_tl", "c_senior", "c_notes", "c_debt", "c_cash",
            "c_net", "c_equity", "c_cap", None, "h_own", "own_sponsor", "own_roll"], 4)
        self.sheet("Operating", [
            "yr", "fy", None, "h_link", "base_year", "da_pct", "nwc_pct", "tax_rate",
            "capex_pct", "fee_years", "fin_fees", None,
            "h_ops", "growth", "revenue", "margin", "ebitda", "da", "ebit", "capex", "dnwc", None,
            "h_tax", "t_ebit", "t_fee", "t_int", "t_inc", "ebt", "tax", "ni", None,
            "h_fcf", "f_ebitda", "f_tax", "f_capex", "f_nwc", "cfads", "f_int", "fcf_post_int",
            "f_mand", "fcf_post_mand", "conv"], 4)
        self.sheet("Debt", [
            "yr", "fy", "lbl", None, "h_link", "base_year", "rev_commit", "rev_rate", "commit_fee",
            "tl_rate", "notes_rate", "pik_flag", "inc_rate", "sweep_pct", "min_cash", "tl_orig", None,
            "h_cash", "cash_open", "fcf_post_mand", "cash_pre", "avail", "avail2", "ecf",
            "cash_close", None,
            "h_rev", "rev_open", "rev_draw", "rev_repay", "rev_close", None,
            "h_tl", "mand_pct", "tl_open", "tl_mand", "tl_sweep", "tl_close", None,
            "h_notes", "notes_open", "notes_pik", "notes_close", None,
            "h_int", "int_rev", "int_commit", "int_tl", "int_notes", "int_notes_cash",
            "int_notes_pik", "gross_int", "int_inc", "net_cash_int", "tot_int_exp", None,
            "h_sum", "total_debt", "senior_debt", "cash_bal", "net_debt", None,
            "h_credit", "k_ebitda", "k_capex", "k_tax", "lev_total", "lev_senior", "lev_net",
            "cov_ebitda", "cov_capex", "fccr"], 4)
        self.sheet("Returns", [
            "h_link", "exit_year", "exit_mult", "entry_mult", "ltm_ebitda", "e0", "sponsor_eq",
            "pool_pct", "nd0", "fees", None,
            "h_exit", "x_ebitda", "x_ev", "x_rev", "x_tl", "x_notes", "x_cash", "x_nd",
            "x_equity", None,
            "h_dist", "d_gain", "d_pool", "d_inv", "d_own", "d_sponsor", "d_roll", "moic", "irr",
            None, "h_bridge", "b_hdr", "b_e0", "b_growth", "b_mult", "b_debt", "b_fees", "b_gain",
            "b_exit", "b_check", "b_pool", "b_invgain", "b_own", "b_sgain", "b_scheck", None,
            "h_cf", "cf_yr", "cf_sponsor", "irr_check", None,
            "h_wf", "w_hdr", "w1", "w2", "w3", "w4", "w5", "w6"], 4)
        self.sheet("Sensitivity", [
            "h_link", "exit_year", "exit_mult", "entry_mult", "step", "ebitda0",
            "txn_pct", "fin_fees", "min_cash", "debt0", "pool", "e0", None, "yr", "ebitda_t",
            "nd_t"], 4)
        self.sheet("Checks", ["hdr"] + [f"c{i}" for i in range(1, 21)] + [None, "master"], 4)

    # --------------------------------------------------------------------- Inputs
    def inputs(self) -> None:
        sh, a, R = "Inputs", self.a, self.r["Inputs"]
        ws = self.ws[sh]
        self.title(sh, "Halcyon Packaging Co. | Inputs",
                   "USD millions unless stated. Blue cells on a cream fill are hard-coded inputs; "
                   "everything else in the workbook is a formula.")
        self.band(sh, R["h_scn"], "Scenario selector")
        self.lab(sh, "scenario", "Active case (choose from dropdown)", bold=True)
        self.inp(sh, f"D{R['scenario']}", a.case.name, "gen", bold=True, align="center")
        ws.data_validation(f"D{R['scenario']}", {
            "validate": "list", "source": f"=$D${R['scn1']}:$D${R['scn3']}",
            "input_title": "Scenario", "input_message": "Pick Base, Management or Downside."})
        self.lab(sh, "scn_idx", "Case index")
        self.one(sh, "scn_idx", f"=MATCH(D{R['scenario']},D{R['scn1']}:D{R['scn3']},0)", "int")
        for i, key in enumerate(SCENARIO_KEYS):
            self.lab(sh, f"scn{i + 1}", f"Case {i + 1} name", indent=1)
            self.inp(sh, f"D{R[f'scn{i + 1}']}", a.cases[key].name, "gen", align="center")

        def block(header: str, items) -> None:
            self.band(sh, R[header], {"h_txn": "Transaction", "h_fin": "Financing",
                                      "h_op": "Operating (all cases)", "h_exit": "Exit and analysis"}[header])
            for key, text, unit, val, num in items:
                self.lab(sh, key, text, unit)
                self.inp(sh, f"D{R[key]}", val, num, align="right" if num == "gen" else None)

        block("h_txn", [
            ("company", "Target company (fictional)", "", a.company, "gen"),
            ("base_year", "Closing fiscal year (LTM period)", "year", a.base_year, "int"),
            ("ltm_revenue", "LTM revenue", "$mm", a.ltm_revenue, "num"),
            ("ltm_ebitda", "LTM EBITDA", "$mm", a.ltm_ebitda, "num"),
            ("entry_mult", "Entry EV / LTM EBITDA", "x", a.entry_multiple, "mult"),
            ("txn_fee_pct", "Transaction fees", "% of EV", a.transaction_fee_pct, "pct2"),
            ("fin_fee_pct", "Financing fees", "% of committed debt", a.financing_fee_pct, "pct2"),
            ("fee_years", "Financing fee amortization period", "years", a.fee_amort_years, "int"),
            ("min_cash", "Minimum cash (funded at close)", "$mm", a.min_cash, "num"),
            ("rollover", "Management rollover equity", "$mm", a.rollover_equity, "num"),
            ("pool_pct", "Management option pool", "% of equity gain", a.option_pool_pct, "pct")])
        block("h_fin", [
            ("rev_commit", "Revolver commitment", "$mm", a.revolver_commitment, "num"),
            ("rev_draw", "Revolver drawn at close", "$mm", a.revolver_draw_at_close, "num"),
            ("tl_x", "Senior term loan", "x LTM EBITDA", a.term_loan_x, "mult"),
            ("notes_x", "Subordinated notes", "x LTM EBITDA", a.sub_notes_x, "mult"),
            ("base_rate", "Base rate", "%", a.base_rate, "pct2"),
            ("rev_spread", "Revolver spread", "%", a.revolver_spread, "pct2"),
            ("tl_spread", "Term loan spread", "%", a.term_spread, "pct2"),
            ("commit_fee", "Revolver commitment fee (undrawn)", "%", a.commitment_fee, "pct2"),
            ("notes_cash", "Sub notes coupon, cash pay", "%", a.notes_cash_rate, "pct2"),
            ("notes_pik_rate", "Sub notes coupon, if PIK", "%", a.notes_pik_rate, "pct2"),
            ("pik_toggle", "Sub notes PIK toggle (1 = PIK, 0 = cash)", "switch", int(a.notes_pik), "int"),
            ("int_income", "Interest income on cash", "%", a.interest_income_rate, "pct2"),
            ("sweep_pct", "Cash sweep", "% of excess cash flow", a.sweep_pct, "pct")])
        ws.data_validation(f"D{R['pik_toggle']}", {"validate": "list", "source": ["0", "1"]})
        self.layout.header_rows.setdefault(sh, []).extend([R["mand_hdr"], R["case_hdr"]])
        for key in ("mand_hdr", "case_hdr"):
            self.t(sh, f"B{R[key]}", "Projection year", color=WHITE, bold=True, bg=NAVY)
            self.t(sh, f"C{R[key]}", "", bg=NAVY)
            self.t(sh, f"D{R[key]}", "", bg=NAVY)
            for t in range(1, 8):
                self.f(sh, f"{YC[t]}{R[key]}", "=COLUMN()-COLUMN($D$1)", "year", color=WHITE,
                       bold=True, bg=NAVY, align="right")
        self.lab(sh, "mand_amort", "Term loan mandatory amortization", "% of original")
        for t in range(1, 8):
            self.inp(sh, f"{YC[t]}{R['mand_amort']}", a.mandatory_amort[t - 1], "pct")
        block("h_op", [
            ("da_pct", "D&A", "% of revenue", a.da_pct, "pct"),
            ("nwc_pct", "Net working capital investment", "% of revenue change", a.nwc_pct, "pct"),
            ("tax_rate", "Cash tax rate", "% of EBT", a.tax_rate, "pct")])
        block("h_exit", [
            ("exit_year", "Exit year (3 to 7)", "year", a.exit_year, "int"),
            ("sens_step", "Sensitivity step for multiples", "x", 0.5, "mult")])
        ws.data_validation(f"D{R['exit_year']}", {"validate": "list", "source": ["3", "4", "5", "6", "7"]})
        self.lab(sh, "exit_mult", "Exit EV / EBITDA (selected case)", "x")
        self.one(sh, "exit_mult", f"=D{R['live_exit']}", "mult")

        self.band(sh, R["h_cases"], "Operating cases")
        for i, key in enumerate(SCENARIO_KEYS):
            c, case = f"c{i}", a.cases[key]
            self.lab(sh, f"{c}_name", "", bold=True)
            self.one(sh, f"{c}_name", f"=D{R[f'scn{i + 1}']}", "gen", bold=True, color=NAVY)
            self.lab(sh, f"{c}_growth", "Revenue growth", "% YoY", indent=1)
            self.lab(sh, f"{c}_margin", "EBITDA margin", "% of revenue", indent=1)
            self.lab(sh, f"{c}_capex", "Capex", "% of revenue", indent=1)
            self.lab(sh, f"{c}_exit", "Exit EV / EBITDA", "x", indent=1)
            for t in range(1, 8):
                self.inp(sh, f"{YC[t]}{R[f'{c}_growth']}", case.revenue_growth[t - 1], "pct")
                self.inp(sh, f"{YC[t]}{R[f'{c}_margin']}", case.ebitda_margin[t - 1], "pct")
            self.inp(sh, f"D{R[f'{c}_capex']}", case.capex_pct, "pct")
            self.inp(sh, f"D{R[f'{c}_exit']}", case.exit_multiple, "mult")
        self.band(sh, R["h_live"], "Selected case (feeds the model)")
        for key, text, unit, num in (("growth", "Revenue growth", "% YoY", "pct"),
                                     ("margin", "EBITDA margin", "% of revenue", "pct")):
            self.lab(sh, f"live_{key}", text, unit)
            for t in range(1, 8):
                c = YC[t]
                refs = ",".join(f"{c}{R[f'c{i}_{key}']}" for i in range(3))
                self.f(sh, f"{c}{R[f'live_{key}']}", f"=CHOOSE($D${R['scn_idx']},{refs})", num)
        for key, text, unit, num in (("capex", "Capex", "% of revenue", "pct"),
                                     ("exit", "Exit EV / EBITDA", "x", "mult")):
            self.lab(sh, f"live_{key}", text, unit)
            refs = ",".join(f"D{R[f'c{i}_{key}']}" for i in range(3))
            self.one(sh, f"live_{key}", f"=CHOOSE($D${R['scn_idx']},{refs})", num)
        ws.freeze_panes(3, 3)

    # ---------------------------------------------------------------- Transaction
    def transaction(self) -> None:
        sh, R = "Transaction", self.r["Transaction"]
        Inp = lambda k: self.X("Inputs", k)
        self.title(sh, "Transaction | purchase price, sources and uses, capitalization",
                   "USD millions. Fees are paid at close; financing fees are capitalized and amortized.")
        self.band(sh, R["h_pp"], "Purchase price and fees")
        d = lambda k: self.A(sh, k)
        self.link_block(sh, [
            ("ltm_ebitda", "LTM EBITDA", f"={Inp('ltm_ebitda')}", "num"),
            ("entry_mult", "Entry EV / LTM EBITDA", f"={Inp('entry_mult')}", "mult"),
            ("ev", "Enterprise value", f"={d('ltm_ebitda')}*{d('entry_mult')}", "num"),
            ("txn_fee_pct", "Transaction fees (% of EV)", f"={Inp('txn_fee_pct')}", "pct2"),
            ("txn_fees", "Transaction fees", f"={d('ev')}*{d('txn_fee_pct')}", "num"),
            ("fin_fee_pct", "Financing fees (% of committed debt)", f"={Inp('fin_fee_pct')}", "pct2"),
            ("rev_commit", "Revolver commitment", f"={Inp('rev_commit')}", "num"),
            ("fin_base", "Committed debt (revolver + term loan + notes)",
             f"={d('rev_commit')}+{d('s_tl')}+{d('s_notes')}", "num"),
            ("fin_fees", "Financing fees", f"={d('fin_base')}*{d('fin_fee_pct')}", "num"),
            ("min_cash", "Minimum cash funded at close", f"={Inp('min_cash')}", "num")])

        for band, hdr, title_ in (("h_s", "s_hdr", "Sources"), ("h_u", "u_hdr", "Uses")):
            self.band(sh, R[band], f"{title_} of funds")
            r = R[hdr]
            for col, txt in zip("DEFG", ("$mm", "% of total", "x LTM EBITDA",
                                         "Rate" if title_ == "Sources" else ""), strict=True):
                self.t(sh, f"{col}{r}", txt, bold=True, align="right", bottom=True)
            self.t(sh, f"B{r}", title_, bold=True, bottom=True)
            self.t(sh, f"C{r}", "", bottom=True)
        src = [("s_rev", "Revolver (drawn at close)", f"={Inp('rev_draw')}"),
               ("s_tl", "Senior term loan", f"={Inp('tl_x')}*{d('ltm_ebitda')}"),
               ("s_notes", "Subordinated notes", f"={Inp('notes_x')}*{d('ltm_ebitda')}"),
               ("s_roll", "Management rollover equity", f"={Inp('rollover')}"),
               ("s_sponsor", "Sponsor equity (plug)",
                f"={d('u_total')}-{d('s_rev')}-{d('s_tl')}-{d('s_notes')}-{d('s_roll')}")]
        for key, text, fml in src:
            self.lab(sh, key, text, indent=1)
            self.one(sh, key, fml, "num")
        rates = {"s_rev": "rev_rate", "s_tl": "tl_rate", "s_notes": "notes_rate"}
        for key, rk in rates.items():
            self.one(sh, key, f"={self.X('Debt', rk)}", "pct2", col="G")
        uses = [("u_ev", "Purchase of enterprise value", f"={d('ev')}"),
                ("u_txn", "Transaction fees", f"={d('txn_fees')}"),
                ("u_fin", "Financing fees", f"={d('fin_fees')}"),
                ("u_cash", "Minimum cash to balance sheet", f"={d('min_cash')}")]
        for key, text, fml in uses:
            self.lab(sh, key, text, indent=1)
            self.one(sh, key, fml, "num")
        for tot, first, last, text in (("s_total", "s_rev", "s_sponsor", "Total sources"),
                                       ("u_total", "u_ev", "u_cash", "Total uses")):
            self.lab(sh, tot, text, bold=True, top=True)
            self.one(sh, tot, f"=SUM({d(first)}:{d(last)})", "num", bold=True, top=True)
            for col in "EFG":
                self.t(sh, f"{col}{R[tot]}", "", top=True)
        for keys, tot in ((["s_rev", "s_tl", "s_notes", "s_roll", "s_sponsor", "s_total"], "s_total"),
                          (["u_ev", "u_txn", "u_fin", "u_cash", "u_total"], "u_total")):
            for k in keys:
                bold, top = k == tot, k == tot
                self.one(sh, k, f"={d(k)}/{self.A(sh, tot, absolute=True)}", "pct", col="E",
                         bold=bold, top=top)
                self.one(sh, k, f"={d(k)}/{d('ltm_ebitda')}", "mult2", col="F", bold=bold, top=top)
        self.lab(sh, "su_check", "Check: total sources less total uses", bold=True)
        self.one(sh, "su_check", f"={d('s_total')}-{d('u_total')}", "num", bold=True)

        self.band(sh, R["h_cap"], "Capitalization at close")
        r = R["cap_hdr"]
        self.t(sh, f"B{r}", "Tranche", bold=True, bottom=True)
        self.t(sh, f"C{r}", "", bottom=True)
        for col, txt in zip("DEF", ("$mm", "% of total cap", "x LTM EBITDA"), strict=True):
            self.t(sh, f"{col}{r}", txt, bold=True, align="right", bottom=True)
        cap = [("c_rev", "Revolver", f"={d('s_rev')}", False),
               ("c_tl", "Senior term loan", f"={d('s_tl')}", False),
               ("c_senior", "Total senior debt", f"={d('c_rev')}+{d('c_tl')}", True),
               ("c_notes", "Subordinated notes", f"={d('s_notes')}", False),
               ("c_debt", "Total debt", f"={d('c_senior')}+{d('c_notes')}", True),
               ("c_cash", "Less: cash", f"={d('u_cash')}", False),
               ("c_net", "Net debt", f"={d('c_debt')}-{d('c_cash')}", True),
               ("c_equity", "Total equity (sponsor + rollover)", f"={d('s_sponsor')}+{d('s_roll')}", False),
               ("c_cap", "Total capitalization (debt + equity)", f"={d('c_debt')}+{d('c_equity')}", True)]
        for key, text, fml, bold in cap:
            self.lab(sh, key, text, bold=bold, indent=0 if bold else 1, top=bold)
            self.one(sh, key, fml, "num", bold=bold, top=bold)
            if key not in ("c_cash", "c_net"):
                self.one(sh, key, f"={d(key)}/{self.A(sh, 'c_cap', absolute=True)}", "pct",
                         col="E", bold=bold, top=bold)
            else:
                self.t(sh, f"E{R[key]}", "", top=bold)
            self.one(sh, key, f"={d(key)}/{d('ltm_ebitda')}", "mult2", col="F", bold=bold, top=bold)
        self.band(sh, R["h_own"], "Ownership at close")
        self.lab(sh, "own_sponsor", "Sponsor", indent=1)
        self.one(sh, "own_sponsor", f"={d('s_sponsor')}/{d('c_equity')}", "pct")
        self.lab(sh, "own_roll", "Management rollover", indent=1)
        self.one(sh, "own_roll", f"={d('s_roll')}/{d('c_equity')}", "pct")
        self.ws[sh].freeze_panes(3, 0)

    # ------------------------------------------------------------------ Operating
    def operating(self) -> None:
        sh, R = "Operating", self.r["Operating"]
        Inp = lambda k: self.X("Inputs", k)
        self.title(sh, "Operating projections and free cash flow",
                   "USD millions. Costs are shown as positive numbers and subtracted. "
                   "Taxes are computed after interest.")
        self.year_header(sh, "yr", "fy")
        self.band(sh, R["h_link"], "Linked assumptions")
        self.link_block(sh, [
            ("base_year", "Closing fiscal year", f"={Inp('base_year')}", "int"),
            ("da_pct", "D&A (% of revenue)", f"={Inp('da_pct')}", "pct"),
            ("nwc_pct", "NWC investment (% of revenue change)", f"={Inp('nwc_pct')}", "pct"),
            ("tax_rate", "Cash tax rate", f"={Inp('tax_rate')}", "pct"),
            ("capex_pct", "Capex (% of revenue, selected case)", f"={Inp('live_capex')}", "pct"),
            ("fee_years", "Financing fee amortization (years)", f"={Inp('fee_years')}", "int"),
            ("fin_fees", "Capitalized financing fees", f"={self.X('Transaction', 'fin_fees')}", "num")])
        d = lambda k: self.A(sh, k, absolute=True)
        self.band(sh, R["h_ops"], "Operating build")
        for k, text, unit in (("growth", "Revenue growth", "%"), ("revenue", "Revenue", "$mm"),
                              ("margin", "EBITDA margin", "%"), ("ebitda", "EBITDA", "$mm"),
                              ("da", "Less: D&A", "$mm"), ("ebit", "EBIT", "$mm"),
                              ("capex", "Capex", "$mm"), ("dnwc", "Increase in net working capital", "$mm")):
            self.lab(sh, k, text, unit, bold=k in ("revenue", "ebitda", "ebit"))
        self.ser(sh, "growth", lambda col, p, t: f"={self.X('Inputs', 'live_growth', col)}", num="pct",
                 italic=True)
        self.one(sh, "revenue", f"={Inp('ltm_revenue')}", bold=True)
        self.ser(sh, "revenue", lambda col, p, t: f"={p}{R['revenue']}*(1+{col}{R['growth']})", bold=True)
        self.one(sh, "margin", f"=D{R['ebitda']}/D{R['revenue']}", "pct", italic=True)
        self.ser(sh, "margin", lambda col, p, t: f"={self.X('Inputs', 'live_margin', col)}", num="pct",
                 italic=True)
        self.one(sh, "ebitda", f"={Inp('ltm_ebitda')}", bold=True)
        self.ser(sh, "ebitda", lambda col, p, t: f"={col}{R['revenue']}*{col}{R['margin']}", bold=True)
        self.ser(sh, "da", lambda col, p, t: f"={d('da_pct')}*{col}{R['revenue']}")
        self.ser(sh, "ebit", lambda col, p, t: f"={col}{R['ebitda']}-{col}{R['da']}", bold=True, top=True)
        self.ser(sh, "capex", lambda col, p, t: f"={d('capex_pct')}*{col}{R['revenue']}")
        self.ser(sh, "dnwc", lambda col, p, t: f"={d('nwc_pct')}*({col}{R['revenue']}-{p}{R['revenue']})")

        self.band(sh, R["h_tax"], "Taxes and net income")
        for k, text in (("t_ebit", "EBIT"), ("t_fee", "Less: financing fee amortization"),
                        ("t_int", "Less: total interest expense (cash + PIK)"),
                        ("t_inc", "Plus: interest income"), ("ebt", "Pre-tax income"),
                        ("tax", "Less: cash taxes"), ("ni", "Net income")):
            self.lab(sh, k, text, bold=k in ("ebt", "ni"))
        self.ser(sh, "t_ebit", lambda col, p, t: f"={col}{R['ebit']}")
        self.ser(sh, "t_fee", lambda col, p, t: f"=IF({col}{R['yr']}<={d('fee_years')},"
                 f"{d('fin_fees')}/{d('fee_years')},0)")
        self.ser(sh, "t_int", lambda col, p, t: f"={self.X('Debt', 'tot_int_exp', col).replace('$', '')}")
        self.ser(sh, "t_inc", lambda col, p, t: f"={self.X('Debt', 'int_inc', col).replace('$', '')}")
        self.ser(sh, "ebt", lambda col, p, t: f"={col}{R['t_ebit']}-{col}{R['t_fee']}-{col}{R['t_int']}"
                 f"+{col}{R['t_inc']}", bold=True, top=True)
        self.ser(sh, "tax", lambda col, p, t: f"=MAX(0,{col}{R['ebt']})*{d('tax_rate')}")
        self.ser(sh, "ni", lambda col, p, t: f"={col}{R['ebt']}-{col}{R['tax']}", bold=True, top=True)

        self.band(sh, R["h_fcf"], "Free cash flow")
        for k, text, bold in (("f_ebitda", "EBITDA", False), ("f_tax", "Less: cash taxes", False),
                              ("f_capex", "Less: capex", False), ("f_nwc", "Less: increase in NWC", False),
                              ("cfads", "Free cash flow before debt service", True),
                              ("f_int", "Less: net cash interest", False),
                              ("fcf_post_int", "Free cash flow after interest", True),
                              ("f_mand", "Less: mandatory term loan amortization", False),
                              ("fcf_post_mand", "Free cash flow after debt service", True),
                              ("conv", "Cash conversion (FCF before debt service / EBITDA)", False)):
            self.lab(sh, k, text, bold=bold)
        self.ser(sh, "f_ebitda", lambda col, p, t: f"={col}{R['ebitda']}")
        self.ser(sh, "f_tax", lambda col, p, t: f"={col}{R['tax']}")
        self.ser(sh, "f_capex", lambda col, p, t: f"={col}{R['capex']}")
        self.ser(sh, "f_nwc", lambda col, p, t: f"={col}{R['dnwc']}")
        self.ser(sh, "cfads", lambda col, p, t: f"={col}{R['f_ebitda']}-{col}{R['f_tax']}-{col}{R['f_capex']}"
                 f"-{col}{R['f_nwc']}", bold=True, top=True)
        self.ser(sh, "f_int", lambda col, p, t: f"={self.X('Debt', 'net_cash_int', col).replace('$', '')}")
        self.ser(sh, "fcf_post_int", lambda col, p, t: f"={col}{R['cfads']}-{col}{R['f_int']}",
                 bold=True, top=True)
        self.ser(sh, "f_mand", lambda col, p, t: f"={self.X('Debt', 'tl_mand', col).replace('$', '')}")
        self.ser(sh, "fcf_post_mand", lambda col, p, t: f"={col}{R['fcf_post_int']}-{col}{R['f_mand']}",
                 bold=True, top=True)
        self.ser(sh, "conv", lambda col, p, t: f"={col}{R['cfads']}/{col}{R['ebitda']}", num="pct",
                 italic=True)
        self.ws[sh].freeze_panes(R["fy"], 3)

    # ----------------------------------------------------------------------- Debt
    def debt(self) -> None:
        sh, R = "Debt", self.r["Debt"]
        Inp = lambda k: self.X("Inputs", k)
        self.title(sh, "Debt schedules, interest and credit statistics",
                   "USD millions. Interest accrues on opening balances (no circular references). "
                   "Cash waterfall: mandatory amortization, revolver, then the cash sweep.")
        self.year_header(sh, "yr", "fy")
        self.t(sh, f"B{R['lbl']}", "Chart label", italic=True, color=GREY)
        self.t(sh, f"D{R['lbl']}", "Close", italic=True, color=GREY, align="right")
        for t in range(1, 8):
            self.f(sh, f"{YC[t]}{R['lbl']}", f'="FY"&{YC[t]}{R["fy"]}&"E"', italic=True,
                   color=GREY, align="right")
        self.band(sh, R["h_link"], "Linked assumptions")
        self.link_block(sh, [
            ("base_year", "Closing fiscal year", f"={Inp('base_year')}", "int"),
            ("rev_commit", "Revolver commitment", f"={Inp('rev_commit')}", "num"),
            ("rev_rate", "Revolver rate (base + spread)", f"={Inp('base_rate')}+{Inp('rev_spread')}", "pct2"),
            ("commit_fee", "Commitment fee on undrawn revolver", f"={Inp('commit_fee')}", "pct2"),
            ("tl_rate", "Term loan rate (base + spread)", f"={Inp('base_rate')}+{Inp('tl_spread')}", "pct2"),
            ("notes_rate", "Sub notes rate (cash or PIK coupon)",
             f"=IF({Inp('pik_toggle')}=1,{Inp('notes_pik_rate')},{Inp('notes_cash')})", "pct2"),
            ("pik_flag", "Sub notes PIK switch", f"={Inp('pik_toggle')}", "int"),
            ("inc_rate", "Interest income rate", f"={Inp('int_income')}", "pct2"),
            ("sweep_pct", "Cash sweep (% of excess cash flow)", f"={Inp('sweep_pct')}", "pct"),
            ("min_cash", "Minimum cash", f"={Inp('min_cash')}", "num"),
            ("tl_orig", "Original term loan", f"={self.X('Transaction', 's_tl')}", "num")])
        d = lambda k: self.A(sh, k, absolute=True)
        T = lambda k: self.X("Transaction", k)

        def line(key, text, fn, t0=1, close=None, **kw):
            self.lab(sh, key, text, bold=kw.get("bold", False))
            if close:
                self.one(sh, key, close, **kw)
            self.ser(sh, key, fn, t0=t0, **kw)

        self.band(sh, R["h_cash"], "Cash waterfall")
        line("cash_open", "Opening cash", lambda c, p, t: f"={p}{R['cash_close']}")
        line("fcf_post_mand", "Free cash flow after debt service",
             lambda c, p, t: f"={self.X('Operating', 'fcf_post_mand', c).replace('$', '')}")
        line("cash_pre", "Cash before revolver and sweep",
             lambda c, p, t: f"={c}{R['cash_open']}+{c}{R['fcf_post_mand']}", top=True)
        line("avail", "Cash above minimum (before revolver)",
             lambda c, p, t: f"={c}{R['cash_pre']}-{d('min_cash')}")
        line("avail2", "Cash above minimum (after revolver)",
             lambda c, p, t: f"={c}{R['avail']}+{c}{R['rev_draw']}-{c}{R['rev_repay']}")
        line("ecf", "Excess cash flow subject to sweep",
             lambda c, p, t: f"=MAX(0,{c}{R['fcf_post_mand']}-{c}{R['rev_repay']})")
        line("cash_close", "Closing cash",
             lambda c, p, t: f"={c}{R['cash_pre']}+{c}{R['rev_draw']}-{c}{R['rev_repay']}-{c}{R['tl_sweep']}",
             close=f"={T('u_cash')}", bold=True, top=True)

        self.band(sh, R["h_rev"], "Revolver")
        line("rev_open", "Opening balance", lambda c, p, t: f"={p}{R['rev_close']}")
        line("rev_draw", "Draw", lambda c, p, t: f"=MIN(MAX(-{c}{R['avail']},0),{d('rev_commit')}-{c}{R['rev_open']})")
        line("rev_repay", "Repayment", lambda c, p, t: f"=MIN({c}{R['rev_open']},MAX({c}{R['avail']},0))")
        line("rev_close", "Closing balance",
             lambda c, p, t: f"={c}{R['rev_open']}+{c}{R['rev_draw']}-{c}{R['rev_repay']}",
             close=f"={T('s_rev')}", bold=True, top=True)

        self.band(sh, R["h_tl"], "Senior term loan")
        line("mand_pct", "Mandatory amortization (% of original)",
             lambda c, p, t: f"={self.X('Inputs', 'mand_amort', c).replace('$', '')}", num="pct")
        line("tl_open", "Opening balance", lambda c, p, t: f"={p}{R['tl_close']}")
        line("tl_mand", "Mandatory amortization",
             lambda c, p, t: f"=MIN({c}{R['tl_open']},{c}{R['mand_pct']}*{d('tl_orig')})")
        line("tl_sweep", "Cash sweep",
             lambda c, p, t: f"=MAX(0,MIN({c}{R['tl_open']}-{c}{R['tl_mand']},"
             f"{d('sweep_pct')}*{c}{R['ecf']},{c}{R['avail2']}))")
        line("tl_close", "Closing balance",
             lambda c, p, t: f"={c}{R['tl_open']}-{c}{R['tl_mand']}-{c}{R['tl_sweep']}",
             close=f"={T('s_tl')}", bold=True, top=True)

        self.band(sh, R["h_notes"], "Subordinated notes")
        line("notes_open", "Opening balance", lambda c, p, t: f"={p}{R['notes_close']}")
        line("notes_pik", "PIK accrual", lambda c, p, t: f"={c}{R['int_notes_pik']}")
        line("notes_close", "Closing balance",
             lambda c, p, t: f"={c}{R['notes_open']}+{c}{R['notes_pik']}",
             close=f"={T('s_notes')}", bold=True, top=True)

        self.band(sh, R["h_int"], "Interest (on opening balances)")
        line("int_rev", "Revolver interest", lambda c, p, t: f"={c}{R['rev_open']}*{d('rev_rate')}")
        line("int_commit", "Commitment fee on undrawn revolver",
             lambda c, p, t: f"=({d('rev_commit')}-{c}{R['rev_open']})*{d('commit_fee')}")
        line("int_tl", "Term loan interest", lambda c, p, t: f"={c}{R['tl_open']}*{d('tl_rate')}")
        line("int_notes", "Sub notes interest (total coupon)",
             lambda c, p, t: f"={c}{R['notes_open']}*{d('notes_rate')}")
        line("int_notes_cash", "of which paid in cash",
             lambda c, p, t: f"=IF({d('pik_flag')}=1,0,{c}{R['int_notes']})")
        line("int_notes_pik", "of which PIK (accrues to principal)",
             lambda c, p, t: f"=IF({d('pik_flag')}=1,{c}{R['int_notes']},0)")
        line("gross_int", "Gross cash interest",
             lambda c, p, t: f"={c}{R['int_rev']}+{c}{R['int_commit']}+{c}{R['int_tl']}+{c}{R['int_notes_cash']}",
             bold=True, top=True)
        line("int_inc", "Interest income on cash", lambda c, p, t: f"={c}{R['cash_open']}*{d('inc_rate')}")
        line("net_cash_int", "Net cash interest", lambda c, p, t: f"={c}{R['gross_int']}-{c}{R['int_inc']}",
             bold=True)
        line("tot_int_exp", "Total interest expense (cash + PIK)",
             lambda c, p, t: f"={c}{R['int_rev']}+{c}{R['int_commit']}+{c}{R['int_tl']}+{c}{R['int_notes']}")

        self.band(sh, R["h_sum"], "Debt summary")
        for key, text, fn, bold in (
                ("total_debt", "Total debt", lambda c: f"={c}{R['rev_close']}+{c}{R['tl_close']}+{c}{R['notes_close']}", True),
                ("senior_debt", "Senior debt (revolver + term loan)", lambda c: f"={c}{R['rev_close']}+{c}{R['tl_close']}", False),
                ("cash_bal", "Cash", lambda c: f"={c}{R['cash_close']}", False),
                ("net_debt", "Net debt", lambda c: f"={c}{R['total_debt']}-{c}{R['cash_bal']}", True)):
            self.lab(sh, key, text, bold=bold)
            for t in range(8):
                self.f(sh, f"{YC[t]}{R[key]}", fn(YC[t]), "num", bold=bold)

        self.band(sh, R["h_credit"], "Credit statistics")
        Op = lambda k, c: f"={self.X('Operating', k, c).replace('$', '')}"
        self.lab(sh, "k_ebitda", "EBITDA")
        for t in range(8):
            self.f(sh, f"{YC[t]}{R['k_ebitda']}", Op("ebitda", YC[t]))
        self.lab(sh, "k_capex", "Capex")
        self.ser(sh, "k_capex", lambda c, p, t: Op("capex", c))
        self.lab(sh, "k_tax", "Cash taxes")
        self.ser(sh, "k_tax", lambda c, p, t: Op("tax", c))
        for key, text, num_key in (("lev_total", "Total debt / EBITDA", "total_debt"),
                                   ("lev_senior", "Senior debt / EBITDA", "senior_debt"),
                                   ("lev_net", "Net debt / EBITDA", "net_debt")):
            self.lab(sh, key, text, bold=key == "lev_total")
            for t in range(8):
                c = YC[t]
                self.f(sh, f"{c}{R[key]}", f"={c}{R[num_key]}/{c}{R['k_ebitda']}", "mult2",
                       bold=key == "lev_total")
        eps = "0.000000001"
        cov = [("cov_ebitda", "Interest coverage (EBITDA / gross cash interest)",
                lambda c: (f"{c}{R['gross_int']}", f"{c}{R['k_ebitda']}")),
               ("cov_capex", "(EBITDA - capex) / gross cash interest",
                lambda c: (f"{c}{R['gross_int']}", f"({c}{R['k_ebitda']}-{c}{R['k_capex']})")),
               ("fccr", "Fixed charge coverage (FCCR)",
                lambda c: (f"({c}{R['gross_int']}+{c}{R['tl_mand']})",
                           f"({c}{R['k_ebitda']}-{c}{R['k_capex']}-{c}{R['k_tax']})"))]
        for key, text, fn in cov:
            self.lab(sh, key, text, bold=key == "cov_ebitda")
            for t in range(1, 8):
                den, num = fn(YC[t])
                self.f(sh, f"{YC[t]}{R[key]}", f'=IF({den}>{eps},{num}/{den},"n.a.")', "mult2",
                       align="right", bold=key == "cov_ebitda")
        self.ws[sh].freeze_panes(R["lbl"], 3)

    # -------------------------------------------------------------------- Returns
    def returns(self) -> None:
        sh, R = "Returns", self.r["Returns"]
        Inp = lambda k: self.X("Inputs", k)
        T = lambda k: self.X("Transaction", k)
        self.title(sh, "Exit, returns and value creation bridge",
                   "USD millions. Sponsor IRR is the closed form MOIC^(1/years) - 1, "
                   "cross-checked against Excel's IRR() on the cash flow row.")
        self.band(sh, R["h_link"], "Linked inputs")
        self.link_block(sh, [
            ("exit_year", "Exit year", f"={Inp('exit_year')}", "int"),
            ("exit_mult", "Exit EV / EBITDA", f"={Inp('exit_mult')}", "mult"),
            ("entry_mult", "Entry EV / EBITDA", f"={Inp('entry_mult')}", "mult"),
            ("ltm_ebitda", "LTM EBITDA at entry", f"={Inp('ltm_ebitda')}", "num"),
            ("e0", "Total entry equity (sponsor + rollover)", f"={T('c_equity')}", "num"),
            ("sponsor_eq", "Sponsor equity invested", f"={T('s_sponsor')}", "num"),
            ("pool_pct", "Option pool (% of equity gain)", f"={Inp('pool_pct')}", "pct"),
            ("nd0", "Net debt at close", f"={T('c_net')}", "num"),
            ("fees", "Transaction and financing fees",
             f"={T('txn_fees')}+{T('fin_fees')}", "num")])
        d = lambda k: self.A(sh, k)

        def idx(sheet, key):
            row = self.r[sheet][key]
            return f"=INDEX({sheet}!$E${row}:$K${row},{d('exit_year')})"

        self.band(sh, R["h_exit"], "Exit valuation")
        self.link_block(sh, [
            ("x_ebitda", "EBITDA in exit year", idx("Operating", "ebitda"), "num"),
            ("x_ev", "Exit enterprise value", f"={d('x_ebitda')}*{d('exit_mult')}", "num"),
            ("x_rev", "Less: revolver", idx("Debt", "rev_close"), "num"),
            ("x_tl", "Less: senior term loan", idx("Debt", "tl_close"), "num"),
            ("x_notes", "Less: subordinated notes (incl. PIK)", idx("Debt", "notes_close"), "num"),
            ("x_cash", "Plus: cash", idx("Debt", "cash_close"), "num"),
            ("x_nd", "Net debt at exit", f"={d('x_rev')}+{d('x_tl')}+{d('x_notes')}-{d('x_cash')}", "num"),
            ("x_equity", "Exit equity value (100%)", f"={d('x_ev')}-{d('x_nd')}", "num")],
            bold=("x_equity",))
        self.band(sh, R["h_dist"], "Distribution and sponsor returns")
        self.link_block(sh, [
            ("d_gain", "Equity gain (exit equity less entry equity)", f"={d('x_equity')}-{d('e0')}", "num"),
            ("d_pool", "Management option pool value", f"={d('pool_pct')}*MAX(0,{d('d_gain')})", "num"),
            ("d_inv", "Equity to investors (sponsor + rollover)", f"={d('x_equity')}-{d('d_pool')}", "num"),
            ("d_own", "Sponsor ownership of investor equity", f"={d('sponsor_eq')}/{d('e0')}", "pct"),
            ("d_sponsor", "Sponsor exit proceeds", f"={d('d_own')}*{d('d_inv')}", "num"),
            ("d_roll", "Rollover holder exit proceeds", f"={d('d_inv')}-{d('d_sponsor')}", "num"),
            ("moic", "Sponsor MOIC", f"={d('d_sponsor')}/{d('sponsor_eq')}", "mult2"),
            ("irr", "Sponsor IRR", f"=IF({d('moic')}<=0,-1,{d('moic')}^(1/{d('exit_year')})-1)", "pct")],
            bold=("moic", "irr"))

        self.band(sh, R["h_bridge"], "Value creation bridge (equity gain, 100% of equity)")
        r = R["b_hdr"]
        self.t(sh, f"B{r}", "Convention: EBITDA growth at the entry multiple; multiple expansion on exit EBITDA",
               italic=True, color=GREY, bottom=True)
        self.t(sh, f"C{r}", "", bottom=True)
        for col, txt in zip("DEF", ("$mm", "% of gain", "x entry equity"), strict=True):
            self.t(sh, f"{col}{r}", txt, bold=True, align="right", bottom=True)
        comp = [("b_e0", "Entry equity (sponsor + rollover)", f"={d('e0')}", False),
                ("b_growth", "EBITDA growth (x entry multiple)",
                 f"={d('entry_mult')}*({d('x_ebitda')}-{d('ltm_ebitda')})", False),
                ("b_mult", "Multiple expansion (x exit EBITDA)",
                 f"=({d('exit_mult')}-{d('entry_mult')})*{d('x_ebitda')}", False),
                ("b_debt", "Net debt paydown (net debt at close less at exit)",
                 f"={d('nd0')}-{d('x_nd')}", False),
                ("b_fees", "Transaction and financing costs", f"=-{d('fees')}", False),
                ("b_gain", "Total equity gain", f"=SUM({d('b_growth')}:{d('b_fees')})", True),
                ("b_exit", "Exit equity value", f"={d('b_e0')}+{d('b_gain')}", True),
                ("b_check", "Check: bridge exit equity less valuation exit equity",
                 f"={d('b_exit')}-{d('x_equity')}", False)]
        for key, text, fml, bold in comp:
            self.lab(sh, key, text, bold=bold, indent=0 if bold else 1, top=bold)
            self.one(sh, key, fml, "num", bold=bold, top=bold)
        for key in ("b_growth", "b_mult", "b_debt", "b_fees", "b_gain"):
            self.one(sh, key, f"=IF(ABS({d('b_gain')})>0.000000001,{d(key)}/{d('b_gain')},0)", "pct",
                     col="E", bold=key == "b_gain", top=key == "b_gain")
        for key in ("b_growth", "b_mult", "b_debt", "b_fees", "b_gain"):
            self.one(sh, key, f"={d(key)}/{d('b_e0')}", "mult2", col="F", bold=key == "b_gain",
                     top=key == "b_gain")
        sp = [("b_pool", "Less: management option pool dilution", f"=-{d('d_pool')}"),
              ("b_invgain", "Gain to investors (sponsor + rollover)", f"={d('b_gain')}+{d('b_pool')}"),
              ("b_own", "Sponsor share", f"={d('d_own')}"),
              ("b_sgain", "Sponsor equity gain", f"={d('b_own')}*{d('b_invgain')}"),
              ("b_scheck", "Check: sponsor gain less (proceeds less equity invested)",
               f"={d('b_sgain')}-({d('d_sponsor')}-{d('sponsor_eq')})")]
        for key, text, fml in sp:
            self.lab(sh, key, text, indent=1, bold=key == "b_sgain")
            self.one(sh, key, fml, "pct" if key == "b_own" else "num", bold=key == "b_sgain")

        self.band(sh, R["h_cf"], "Sponsor cash flows")
        self.layout.header_rows.setdefault(sh, []).append(R["cf_yr"])
        self.t(sh, f"B{R['cf_yr']}", "Year", color=WHITE, bold=True, bg=NAVY)
        self.t(sh, f"C{R['cf_yr']}", "", bg=NAVY)
        for t in range(8):
            self.f(sh, f"{YC[t]}{R['cf_yr']}", "=COLUMN()-COLUMN($D$1)", "year", color=WHITE,
                   bold=True, bg=NAVY, align="right")
        self.lab(sh, "cf_sponsor", "Sponsor equity cash flow", "$mm")
        self.one(sh, "cf_sponsor", f"=-{self.A(sh, 'sponsor_eq', absolute=True)}", "num")
        self.ser(sh, "cf_sponsor", lambda c, p, t: f"=IF({c}{R['cf_yr']}=$D${R['exit_year']},$D${R['d_sponsor']},0)")
        self.lab(sh, "irr_check", "Sponsor IRR using Excel IRR() on the row above", bold=True)
        self.one(sh, "irr_check", f"=IRR(D{R['cf_sponsor']}:K{R['cf_sponsor']})", "pct", bold=True)

        self.band(sh, R["h_wf"], "Bridge chart data (waterfall helper)")
        r = R["w_hdr"]
        for col, txt in zip("BDEFGH", ("Step", "Base (hidden)", "Equity", "Increase", "Decrease", "Level after"),
                            strict=True):
            self.t(sh, f"{col}{r}", txt, bold=True, bottom=True, align=None if col == "B" else "right")
        steps = [("w1", "Entry equity", None), ("w2", "EBITDA growth", "b_growth"),
                 ("w3", "Multiple expansion", "b_mult"), ("w4", "Net debt paydown", "b_debt"),
                 ("w5", "Transaction costs", "b_fees"), ("w6", "Exit equity", None)]
        for i, (key, text, src) in enumerate(steps):
            row = R[key]
            self.t(sh, f"B{row}", text)
            if src is None:
                level = d("b_e0") if i == 0 else d("b_exit")
                self.f(sh, f"H{row}", f"={level}", "num")
                self.f(sh, f"E{row}", f"=H{row}", "num")
            else:
                prev = R[steps[i - 1][0]]
                self.f(sh, f"H{row}", f"=H{prev}+{d(src)}", "num")
                self.f(sh, f"D{row}", f"=MIN(H{prev},H{row})", "num")
                self.f(sh, f"F{row}", f"=MAX({d(src)},0)", "num")
                self.f(sh, f"G{row}", f"=MAX(-{d(src)},0)", "num")
        self.ws[sh].freeze_panes(3, 3)

    # ---------------------------------------------------------------- Sensitivity
    def sensitivity(self) -> None:
        sh, R = "Sensitivity", self.r["Sensitivity"]
        ws = self.ws[sh]
        Inp = lambda k: self.X("Inputs", k)
        T = lambda k: self.X("Transaction", k)
        self.title(sh, "Sensitivity analysis (live formulas, no data tables)",
                   "Entry multiple, exit multiple and exit year leave the debt schedule unchanged, so each grid "
                   "cell is a closed-form formula on the model rows below.")
        self.band(sh, R["h_link"], "Linked inputs")
        self.year_header_simple(sh)
        self.link_block(sh, [
            ("exit_year", "Exit year", f"={Inp('exit_year')}", "int"),
            ("exit_mult", "Exit EV / EBITDA", f"={Inp('exit_mult')}", "mult"),
            ("entry_mult", "Entry EV / EBITDA", f"={Inp('entry_mult')}", "mult"),
            ("step", "Multiple step", f"={Inp('sens_step')}", "mult"),
            ("ebitda0", "LTM EBITDA", f"={Inp('ltm_ebitda')}", "num"),
            ("txn_pct", "Transaction fees (% of EV)", f"={Inp('txn_fee_pct')}", "pct2"),
            ("fin_fees", "Financing fees", f"={T('fin_fees')}", "num"),
            ("min_cash", "Minimum cash", f"={Inp('min_cash')}", "num"),
            ("debt0", "Total debt at close", f"={T('c_debt')}", "num"),
            ("pool", "Option pool (% of equity gain)", f"={Inp('pool_pct')}", "pct"),
            ("e0", "Total entry equity", f"={T('c_equity')}", "num")])
        self.lab(sh, "ebitda_t", "EBITDA by year")
        self.ser(sh, "ebitda_t", lambda c, p, t: f"={self.X('Operating', 'ebitda', c).replace('$', '')}")
        self.lab(sh, "nd_t", "Net debt by year")
        self.ser(sh, "nd_t", lambda c, p, t: f"={self.X('Debt', 'net_debt', c).replace('$', '')}")
        d = lambda k: self.A(sh, k, absolute=True)
        ebitda_rng = f"$E${R['ebitda_t']}:$K${R['ebitda_t']}"
        nd_rng = f"$E${R['nd_t']}:$K${R['nd_t']}"
        steps = (-2, -1, 0, 1, 2)
        cols = "DEFGH"
        row = R["nd_t"] + 2
        base_fmt = self.wb.add_format({"bold": True, "border": 2, "border_color": NAVY})

        def axis(k: int, centre: str) -> str:
            return f"={d(centre)}+({k})*{d('step')}"

        def grid_title(r: int, text: str) -> None:
            self.band(sh, r, text, last="I")

        def corner(r: int, text: str) -> None:
            self.t(sh, f"C{r}", text, bold=True, bottom=True, align="right")

        def colour(first: int, last: int, three: bool = True) -> None:
            ws.conditional_format(f"D{first}:H{last}", {
                "type": "3_color_scale", "min_color": "#F8CBAD", "mid_color": "#FFFFFF",
                "max_color": "#C6E0B4"})

        grids: dict[str, dict[str, object]] = {}
        moic_first = 0
        # grid 1: entry (rows) x exit (columns)
        grid_title(row, "Grid 1: entry multiple (rows) vs exit multiple (columns), exit year per Inputs")
        for metric in ("moic", "irr"):
            hdr = row + 1 if metric == "moic" else row + 12
            if metric == "irr":
                grid_title(row + 11, "Grid 1: sponsor IRR")
            self.t(sh, f"B{hdr}", "Sponsor MOIC" if metric == "moic" else "Sponsor IRR", bold=True)
            corner(hdr, "Entry \\ Exit")
            for j, k in enumerate(steps):
                if metric == "moic":
                    self.f(sh, f"{cols[j]}{hdr}", axis(k, "exit_mult"), "mult", bold=True, bottom=True)
                else:
                    self.f(sh, f"{cols[j]}{hdr}", f"={cols[j]}{row + 1}", "mult", bold=True, bottom=True)
            if metric == "moic":
                helper = hdr + 1
                self.t(sh, f"B{helper}", "Exit equity at that exit multiple (100%)", italic=True, color=GREY)
                for j in range(5):
                    c = cols[j]
                    self.f(sh, f"{c}{helper}", f"={c}{hdr}*INDEX({ebitda_rng},{d('exit_year')})"
                           f"-INDEX({nd_rng},{d('exit_year')})", "num", italic=True)
                self.t(sh, f"I{hdr}", "Entry equity", italic=True, color=GREY, align="right")
            first = hdr + 2 if metric == "moic" else hdr + 1
            for i, k in enumerate(steps):
                r = first + i
                if metric == "moic":
                    self.f(sh, f"C{r}", axis(k, "entry_mult"), "mult", bold=True)
                    self.f(sh, f"I{r}", f"=C{r}*{d('ebitda0')}*(1+{d('txn_pct')})+{d('fin_fees')}"
                           f"+{d('min_cash')}-{d('debt0')}", "num", italic=True)
                    for j in range(5):
                        c = cols[j]
                        x = f"{c}${helper}"
                        self.f(sh, f"{c}{r}", f"=({x}-{d('pool')}*MAX(0,{x}-$I{r}))/$I{r}", "mult2",
                               bold=(i == 2 and j == 2))
                else:
                    self.f(sh, f"C{r}", f"=C{moic_first + i}", "mult", bold=True)
                    for j in range(5):
                        c = cols[j]
                        m = f"{c}{moic_first + i}"
                        self.f(sh, f"{c}{r}", f"=IF({m}<=0,-1,{m}^(1/{d('exit_year')})-1)", "pct")
            if metric == "moic":
                moic_first = first
            colour(first, first + 4)
            grids.setdefault("entry_exit", {})[metric] = {"first_row": first, "cols": cols,
                                                          "hdr_row": row + 1}
            ws.conditional_format(f"D{first}:H{first + 4}", {
                "type": "formula", "criteria": f"=AND(ROW()={first + 2},COLUMN()=6)", "format": base_fmt})
        grids["entry_exit"]["row_axis"] = "C"
        row = row + 23
        # grid 2: exit multiple (rows) x exit year (columns)
        grid_title(row, "Grid 2: exit multiple (rows) vs exit year (columns), entry multiple per Inputs")
        moic2_first = None
        for metric in ("moic", "irr"):
            hdr = row + 1 if metric == "moic" else row + 9
            if metric == "irr":
                grid_title(row + 8, "Grid 2: sponsor IRR")
            self.t(sh, f"B{hdr}", "Sponsor MOIC" if metric == "moic" else "Sponsor IRR", bold=True)
            corner(hdr, "Exit x \\ Year")
            for j in range(5):
                if metric == "moic":
                    self.inp(sh, f"{cols[j]}{hdr}", 3 + j, "int", bold=True)
                else:
                    self.f(sh, f"{cols[j]}{hdr}", f"={cols[j]}{row + 1}", "int", bold=True, bottom=True)
            first = hdr + 1
            for i, k in enumerate(steps):
                r = first + i
                if metric == "moic":
                    self.f(sh, f"C{r}", axis(k, "exit_mult"), "mult", bold=True)
                    for j in range(5):
                        c = cols[j]
                        x = (f"($C{r}*INDEX({ebitda_rng},{c}${hdr})-INDEX({nd_rng},{c}${hdr}))")
                        self.f(sh, f"{c}{r}", f"=({x}-{d('pool')}*MAX(0,{x}-{d('e0')}))/{d('e0')}", "mult2")
                else:
                    self.f(sh, f"C{r}", f"=C{moic2_first + i}", "mult", bold=True)
                    for j in range(5):
                        c = cols[j]
                        m = f"{c}{moic2_first + i}"
                        self.f(sh, f"{c}{r}", f"=IF({m}<=0,-1,{m}^(1/{c}${hdr})-1)", "pct")
            if metric == "moic":
                moic2_first = first
            colour(first, first + 4)
            grids.setdefault("exit_year_exit_mult", {})[metric] = {"first_row": first, "cols": cols,
                                                                   "hdr_row": hdr}
            ws.conditional_format(f"D{first}:H{first + 4}", {
                "type": "formula",
                "criteria": f"=AND(ROW()={first + 2},D${hdr}={d('exit_year')})", "format": base_fmt})
        note = row + 16
        self.t(sh, f"B{note}", "Leverage vs exit year changes the debt schedule, so it is solved by the "
               "Python engine (see `lbo sens` and the README).", italic=True, color=GREY)
        self.layout.grids = grids
        ws.freeze_panes(3, 3)

    def year_header_simple(self, sh: str) -> None:
        ry = self.r[sh]["yr"]
        self.layout.header_rows.setdefault(sh, []).append(ry)
        self.t(sh, f"B{ry}", "Projection year", color=WHITE, bold=True, bg=NAVY)
        self.t(sh, f"C{ry}", "", bg=NAVY)
        for t in range(8):
            self.f(sh, f"{YC[t]}{ry}", "=COLUMN()-COLUMN($D$1)", "year", color=WHITE, bold=True,
                   bg=NAVY, align="right")

    # --------------------------------------------------------------------- Checks
    def checks(self) -> None:
        sh, R = "Checks", self.r["Checks"]
        ws = self.ws[sh]
        self.title(sh, "Model integrity checks",
                   "Each check reports a violation size; it passes when the size is within tolerance.")
        r = R["hdr"]
        for col, txt in zip("BDEF", ("Check", "Violation", "Tolerance", "Status"), strict=True):
            self.t(sh, f"{col}{r}", txt, color=WHITE, bold=True, bg=NAVY, align=None if col == "B" else "right")
        self.t(sh, f"C{r}", "", bg=NAVY)
        def D(k, c="D"):
            return self.X("Debt", k, c)

        def T(k, c="D"):
            return self.X("Transaction", k, c)

        def Rt(k, c="D"):
            return self.X("Returns", k, c)

        def rng(sheet, key, c0="D", c1="K"):
            row = self.r[sheet][key]
            return f"{sheet}!${c0}${row}:${c1}${row}"

        def tot(sheet, key):
            return f"SUM({rng(sheet, key, 'E')})"

        g1 = self.layout.grids  # filled by sensitivity()
        e1m, e1i = g1["entry_exit"]["moic"], g1["entry_exit"]["irr"]
        e2m, e2i = g1["exit_year_exit_mult"]["moic"], g1["exit_year_exit_mult"]["irr"]
        yr = self.X("Inputs", "exit_year")

        def centre2(g):
            row, hdr = g["first_row"] + 2, g["hdr_row"]
            return (f"INDEX(Sensitivity!$D${row}:$H${row},MATCH({yr},Sensitivity!$D${hdr}:$H${hdr},0))")

        checks = [
            ("Sources equal uses", f"=ABS({T('su_check')})"),
            ("Debt plus equity equals total uses", f"=ABS({T('c_cap')}-{T('u_total')})"),
            ("Debt balances never negative (revolver, term loan, notes)",
             f"=MAX(0,-MIN({rng('Debt', 'rev_close')},{rng('Debt', 'tl_close')},{rng('Debt', 'notes_close')}))"),
            ("Cash never below minimum",
             f"=MAX(0,{D('min_cash')}-MIN({rng('Debt', 'cash_close')}))"),
            ("Revolver never above commitment",
             f"=MAX(0,MAX({rng('Debt', 'rev_close')})-{D('rev_commit')})"),
            ("Revolver roll-forward (closing = opening + draws - repayments)",
             f"=ABS({D('rev_close', 'K')}-({D('rev_close')}+{tot('Debt', 'rev_draw')}-{tot('Debt', 'rev_repay')}))"),
            ("Term loan roll-forward (closing = opening - amortization - sweep)",
             f"=ABS({D('tl_close', 'K')}-({D('tl_close')}-{tot('Debt', 'tl_mand')}-{tot('Debt', 'tl_sweep')}))"),
            ("Sub notes roll-forward (closing = opening + PIK)",
             f"=ABS({D('notes_close', 'K')}-({D('notes_close')}+{tot('Debt', 'notes_pik')}))"),
            ("Cash roll-forward (free cash flow, revolver, sweep)",
             f"=ABS({D('cash_close', 'K')}-({D('cash_close')}+{tot('Debt', 'fcf_post_mand')}"
             f"+{tot('Debt', 'rev_draw')}-{tot('Debt', 'rev_repay')}-{tot('Debt', 'tl_sweep')}))"),
            ("Net debt roll-forward (opening - FCF after interest + PIK)",
             f"=ABS({D('net_debt', 'K')}-({D('net_debt')}-{tot('Operating', 'fcf_post_int')}"
             f"+{tot('Debt', 'int_notes_pik')}))"),
            ("Value creation bridge sums to equity gain", f"=ABS({Rt('b_check')})"),
            ("Sponsor bridge reconciles to proceeds less equity", f"=ABS({Rt('b_scheck')})"),
            ("Closed-form IRR equals Excel IRR()", f"=ABS({Rt('irr')}-{Rt('irr_check')})"),
            ("Grid 1 centre MOIC equals model MOIC",
             f"=ABS(Sensitivity!$F${e1m['first_row'] + 2}-{Rt('moic')})"),
            ("Grid 1 centre IRR equals model IRR",
             f"=ABS(Sensitivity!$F${e1i['first_row'] + 2}-{Rt('irr')})"),
            ("Grid 2 centre MOIC equals model MOIC", f"=ABS({centre2(e2m)}-{Rt('moic')})"),
            ("Grid 2 centre IRR equals model IRR", f"=ABS({centre2(e2i)}-{Rt('irr')})"),
            ("Exit year is an integer between 3 and 7",
             f"=IF(AND({yr}>=3,{yr}<=7,{yr}=INT({yr})),0,1)"),
            ("Revolver drawn at close within commitment",
             f"=MAX(0,{self.X('Inputs', 'rev_draw')}-{self.X('Inputs', 'rev_commit')})"),
            ("Cash taxes never negative", f"=MAX(0,-MIN({rng('Operating', 'tax', 'E')}))"),
        ]
        for i, (text, fml) in enumerate(checks, start=1):
            row = R[f"c{i}"]
            self.t(sh, f"B{row}", text)
            self.f(sh, f"D{row}", fml, "sci")
            self.inp(sh, f"E{row}", 0.000001, "sci")
            self.f(sh, f"F{row}", f'=IF(ISNUMBER(D{row}),IF(D{row}<=E{row},"OK","ERROR"),"ERROR")',
                   "gen", bold=True, align="right")
        first, last = R["c1"], R["c20"]
        row = R["master"]
        self.t(sh, f"B{row}", "Overall model status", bold=True, top=True)
        self.t(sh, f"C{row}", "", top=True)
        self.f(sh, f"D{row}", f'=COUNTIF(F{first}:F{last},"ERROR")', "int", bold=True, top=True)
        self.t(sh, f"E{row}", "errors", italic=True, color=GREY, top=True)
        self.f(sh, f"F{row}", f'=IF(D{row}=0,"ALL CHECKS PASS","CHECK FAILURES")', "gen", bold=True,
               top=True, align="right")
        rngs = f"F{first}:F{last}"
        ws.conditional_format(rngs, {"type": "cell", "criteria": "==", "value": '"OK"',
                                     "format": self.wb.add_format({"bg_color": GOOD})})
        ws.conditional_format(rngs, {"type": "cell", "criteria": "==", "value": '"ERROR"',
                                     "format": self.wb.add_format({"bg_color": BAD})})
        ws.set_column("B:B", 62)
        ws.set_column("F:F", 20)
        ws.freeze_panes(R["hdr"], 3)

    # -------------------------------------------------------------------- Summary
    def summary(self) -> None:
        sh = "Summary"
        ws = self.ws[sh]
        T = lambda k, c="D": self.X("Transaction", k, c)
        Rt = lambda k, c="D": self.X("Returns", k, c)
        Inp = lambda k: self.X("Inputs", k)
        ws.set_column("B:B", 34)
        ws.set_column("C:C", 9)
        self.f(sh, "B1", f'={Inp("company")}&" | LBO returns summary | "&{Inp("scenario")}&" case"',
               "gen", bold=True, size=15, color=NAVY)
        self.t(sh, "B2", "USD millions. Fictional company, illustrative only. All figures link to the model sheets.",
               italic=True, color=GREY)

        def block(r: int, label_col: str, value_col: str, title_: str, items, last: str) -> None:
            fm = self.fmt(color=WHITE, bold=True, bg=NAVY)
            for ci in range(ord(label_col), ord(last) + 1):
                ws.write_blank(f"{chr(ci)}{r}", None, fm)
            self.t(sh, f"{label_col}{r}", title_, color=WHITE, bold=True, bg=NAVY)
            for i, (text, fml, num, bold) in enumerate(items, start=1):
                self.t(sh, f"{label_col}{r + i}", text, bold=bold)
                self.f(sh, f"{value_col}{r + i}", fml, num, bold=bold, align="right")

        block(4, "B", "D", "Transaction", [
            ("Enterprise value", f"={T('ev')}", "num", False),
            ("Entry EV / LTM EBITDA", f"={T('entry_mult')}", "mult", False),
            ("Total debt / LTM EBITDA at close", f"={T('c_debt', 'F')}", "mult2", False),
            ("Sponsor equity", f"={T('s_sponsor')}", "num", False),
            ("Total equity % of capitalization", f"={T('c_equity', 'E')}", "pct", False),
            ("Fees (transaction + financing)", f"={T('txn_fees')}+{T('fin_fees')}", "num", False),
            ("Model checks", f"={self.X('Checks', 'master', 'F')}", "gen", True)], "E")
        block(4, "F", "K", "Returns", [
            ("Exit year", f"={Rt('exit_year')}", "int", False),
            ("Exit EV / EBITDA", f"={Rt('exit_mult')}", "mult", False),
            ("Exit enterprise value", f"={Rt('x_ev')}", "num", False),
            ("Exit equity value (100%)", f"={Rt('x_equity')}", "num", False),
            ("Sponsor proceeds", f"={Rt('d_sponsor')}", "num", False),
            ("Sponsor MOIC", f"={Rt('moic')}", "mult2", True),
            ("Sponsor IRR", f"={Rt('irr')}", "pct", True)], "K")

        fm = self.fmt(color=WHITE, bold=True, bg=NAVY)
        for ci in range(ord("B"), ord("E") + 1):
            ws.write_blank(f"{chr(ci)}13", None, fm)
        for ci in range(ord("F"), ord("K") + 1):
            ws.write_blank(f"{chr(ci)}13", None, fm)
        self.t(sh, "B13", "Sources and uses", color=WHITE, bold=True, bg=NAVY)
        self.t(sh, "F13", "Value creation bridge (100% of equity)", color=WHITE, bold=True, bg=NAVY)
        su = [("Revolver", "s_rev"), ("Senior term loan", "s_tl"), ("Subordinated notes", "s_notes"),
              ("Management rollover", "s_roll"), ("Sponsor equity", "s_sponsor"), ("Total sources", "s_total"),
              ("Enterprise value", "u_ev"), ("Fees and expenses", None), ("Minimum cash", "u_cash"),
              ("Total uses", "u_total")]
        for i, (text, key) in enumerate(su):
            r = 14 + i
            bold = text.startswith("Total")
            self.t(sh, f"B{r}", text, bold=bold, top=bold)
            self.t(sh, f"C{r}", "", top=bold)
            fml = f"={T(key)}" if key else f"={T('u_txn')}+{T('u_fin')}"
            self.f(sh, f"D{r}", fml, "num", bold=bold, top=bold)
            if key:
                tot_key = "s_total" if key.startswith("s_") else "u_total"
                self.f(sh, f"E{r}", f"={T(key)}/{T(tot_key)}", "pct", bold=bold, top=bold)
            else:
                self.f(sh, f"E{r}", f"=({T('u_txn')}+{T('u_fin')})/{T('u_total')}", "pct")
        br = [("Entry equity", "b_e0", False), ("EBITDA growth", "b_growth", False),
              ("Multiple expansion", "b_mult", False), ("Net debt paydown", "b_debt", False),
              ("Transaction costs", "b_fees", False), ("Exit equity value", "b_exit", True),
              ("Less: option pool", "b_pool", False), ("Sponsor equity gain", "b_sgain", True)]
        for i, (text, key, bold) in enumerate(br):
            r = 14 + i
            self.t(sh, f"F{r}", text, bold=bold, top=bold)
            for col in "GHI":
                self.t(sh, f"{col}{r}", "", top=bold)
            self.f(sh, f"J{r}", f"={Rt(key)}", "num", bold=bold, top=bold)
            if key in ("b_growth", "b_mult", "b_debt", "b_fees"):
                self.f(sh, f"K{r}", f"={Rt(key, 'E')}", "pct")
            else:
                self.t(sh, f"K{r}", "", top=bold)
        self.t(sh, "F22", "Sponsor MOIC / IRR", bold=False)
        self.f(sh, "J22", f"={Rt('moic')}", "mult2", bold=True)
        self.f(sh, "K22", f"={Rt('irr')}", "pct", bold=True)

        # time series table
        r0 = 25
        for ci in range(ord("B"), ord("K") + 1):
            ws.write_blank(f"{chr(ci)}{r0}", None, fm)
        self.t(sh, f"B{r0}", "Projections and credit statistics", color=WHITE, bold=True, bg=NAVY)
        for t in range(8):
            self.f(sh, f"{YC[t]}{r0}", f"={self.X('Debt', 'lbl', YC[t]).replace('$', '')}", "gen",
                   color=WHITE, bold=True, bg=NAVY, align="right")
        series = [("Revenue", "Operating", "revenue", "num", 0, False),
                  ("EBITDA", "Operating", "ebitda", "num", 0, True),
                  ("EBITDA margin", "Operating", "margin", "pct", 0, False),
                  ("FCF before debt service", "Operating", "cfads", "num", 1, False),
                  ("FCF after debt service", "Operating", "fcf_post_mand", "num", 1, True),
                  ("Total debt", "Debt", "total_debt", "num", 0, False),
                  ("Net debt", "Debt", "net_debt", "num", 0, False),
                  ("Total debt / EBITDA", "Debt", "lev_total", "mult2", 0, True),
                  ("Senior debt / EBITDA", "Debt", "lev_senior", "mult2", 0, False),
                  ("Interest coverage (EBITDA / cash interest)", "Debt", "cov_ebitda", "mult2", 1, False),
                  ("Fixed charge coverage", "Debt", "fccr", "mult2", 1, False)]
        for i, (text, s, key, num, t0, bold) in enumerate(series, start=1):
            r = r0 + i
            self.t(sh, f"B{r}", text, bold=bold)
            for t in range(t0, 8):
                self.f(sh, f"{YC[t]}{r}", f"={self.X(s, key, YC[t]).replace('$', '')}", num, bold=bold,
                       align="right")
        self.summary_charts(r0 + len(series) + 2)
        ws.freeze_panes(2, 0)
        ws.set_paper(1)
        ws.fit_to_pages(1, 1)
        ws.print_area(f"A1:K{r0 + len(series) + 2 + 34}")

    def summary_charts(self, top: int) -> None:
        sh, wb = "Summary", self.wb
        Rr, Rd = self.r["Returns"], self.r["Debt"]
        font = {"name": "Calibri", "size": 9}
        cats = f"=Returns!$B${Rr['w1']}:$B${Rr['w6']}"
        ch = wb.add_chart({"type": "column", "subtype": "stacked"})
        def labels(color: str, num_format: str) -> dict:
            return {"value": True, "num_format": num_format,
                    "font": {"name": "Calibri", "size": 9, "color": color, "bold": True}}

        for col, name, fill, lab in (("D", "Base", None, None),
                                     ("E", "Equity", NAVY, labels(WHITE, "#,##0;(#,##0);;")),
                                     ("F", "Increase", "#3C8D5A", labels(WHITE, "#,##0;(#,##0);;")),
                                     ("G", "Decrease", "#C0392B", labels(BLACK, "(#,##0);(#,##0);;"))):
            s = {"name": name, "categories": cats, "values": f"=Returns!${col}${Rr['w1']}:${col}${Rr['w6']}",
                 "gap": 40}
            if fill is None:
                s["fill"], s["border"] = {"none": True}, {"none": True}
            else:
                s["fill"] = {"color": fill}
            if lab:
                s["data_labels"] = lab
            ch.add_series(s)
        ch.set_title({"name": "Value creation bridge ($mm, 100% of equity)", "name_font": {**font, "size": 11,
                      "bold": True, "color": NAVY}})
        ch.set_legend({"none": True})
        ch.set_y_axis({"num_format": "#,##0", "num_font": font, "major_gridlines": {"visible": True,
                       "line": {"color": "#D9D9D9"}}})
        ch.set_x_axis({"num_font": font})
        ch.set_size({"width": 480, "height": 290})
        self.ws[sh].insert_chart(f"B{top}", ch, {"x_offset": 4, "y_offset": 4})

        lc = f"=Debt!$D${Rd['lbl']}:$K${Rd['lbl']}"
        col = wb.add_chart({"type": "column"})
        for key, name, color in (("lev_total", "Total debt / EBITDA", NAVY),
                                 ("lev_senior", "Senior debt / EBITDA", "#8FAADC")):
            col.add_series({"name": name, "categories": lc,
                            "values": f"=Debt!$D${Rd[key]}:$K${Rd[key]}", "fill": {"color": color}, "gap": 80})
        line = wb.add_chart({"type": "line"})
        line.add_series({"name": "Interest coverage (right axis)", "categories": lc,
                         "values": f"=Debt!$D${Rd['cov_ebitda']}:$K${Rd['cov_ebitda']}", "y2_axis": True,
                         "line": {"color": "#E9A23B", "width": 2.25},
                         "marker": {"type": "circle", "size": 5, "fill": {"color": "#E9A23B"},
                                    "border": {"color": "#E9A23B"}}})
        col.combine(line)
        col.set_title({"name": "Leverage and coverage", "name_font": {**font, "size": 11, "bold": True,
                       "color": NAVY}})
        col.set_y_axis({"num_format": '0.0"x"', "num_font": font, "major_gridlines": {
            "visible": True, "line": {"color": "#D9D9D9"}}})
        line.set_y2_axis({"num_format": '0.0"x"', "num_font": font})
        col.set_x_axis({"num_font": font})
        col.set_legend({"position": "bottom", "font": font})
        col.set_size({"width": 530, "height": 290})
        self.ws[sh].insert_chart(f"F{top}", col, {"x_offset": 4, "y_offset": 4})

        st = wb.add_chart({"type": "column", "subtype": "stacked"})
        for key, name, color in (("tl_close", "Senior term loan", NAVY),
                                 ("notes_close", "Subordinated notes", "#8FAADC"),
                                 ("rev_close", "Revolver", "#E9A23B")):
            st.add_series({"name": name, "categories": lc, "values": f"=Debt!$D${Rd[key]}:$K${Rd[key]}",
                           "fill": {"color": color}, "gap": 60})
        st.set_title({"name": "Debt paydown by tranche ($mm)", "name_font": {**font, "size": 11,
                      "bold": True, "color": NAVY}})
        st.set_y_axis({"num_format": "#,##0", "num_font": font, "major_gridlines": {
            "visible": True, "line": {"color": "#D9D9D9"}}})
        st.set_x_axis({"num_font": font})
        st.set_legend({"position": "bottom", "font": font})
        st.set_size({"width": 480, "height": 290})
        self.ws[sh].insert_chart(f"B{top + 16}", st, {"x_offset": 4, "y_offset": 4})

        # financing summary next to the third chart
        r = top + 16
        ws = self.ws[sh]
        fm = self.fmt(color=WHITE, bold=True, bg=NAVY)
        for ci in range(ord("F"), ord("K") + 1):
            ws.write_blank(f"{chr(ci)}{r}", None, fm)
        self.t(sh, f"F{r}", "Financing summary", color=WHITE, bold=True, bg=NAVY)
        for c, txt in (("H", "$mm"), ("I", "x EBITDA"), ("J", "Rate"), ("K", "Amort.")):
            self.t(sh, f"{c}{r}", txt, color=WHITE, bold=True, bg=NAVY, align="right")
        T = lambda k, c="D": self.X("Transaction", k, c)
        rows = [("Revolver (commitment)", self.X("Inputs", "rev_commit"), None, T("s_rev", "G"), None),
                ("Senior term loan", T("s_tl"), T("s_tl", "F"), T("s_tl", "G"), self.X("Inputs", "mand_amort", "E")),
                ("Subordinated notes", T("s_notes"), T("s_notes", "F"), T("s_notes", "G"), None)]
        for i, (text, amt, mult, rate, am) in enumerate(rows, start=1):
            self.t(sh, f"F{r + i}", text)
            self.f(sh, f"H{r + i}", f"={amt}", "num")
            if mult:
                self.f(sh, f"I{r + i}", f"={mult}", "mult2")
            self.f(sh, f"J{r + i}", f"={rate}", "pct2")
            if am:
                self.f(sh, f"K{r + i}", f"={am}", "pct")
        self.t(sh, f"F{r + 4}", "Cash sweep (% of excess cash flow)")
        self.f(sh, f"K{r + 4}", f"={self.X('Inputs', 'sweep_pct')}", "pct")
        self.t(sh, f"F{r + 5}", "Sub notes pay in kind (1 = yes)")
        self.f(sh, f"K{r + 5}", f"={self.X('Inputs', 'pik_toggle')}", "int")
        self.t(sh, f"F{r + 6}", "Minimum cash")
        self.f(sh, f"K{r + 6}", f"={self.X('Inputs', 'min_cash')}", "num")
        self.t(sh, f"F{r + 8}", "Interest is charged on opening balances; no circular references.",
               italic=True, color=GREY)

    # ---------------------------------------------------------------------- Cover
    def cover(self) -> None:
        sh = "Cover"
        ws = self.ws[sh]
        ws.set_column("B:B", 30)
        ws.set_column("C:C", 90)
        self.t(sh, "B2", "Halcyon Packaging Co.", bold=True, size=22, color=NAVY)
        self.t(sh, "B3", "Leveraged buyout model", size=14, color=STEEL)
        self.t(sh, "B4", "Fictional company. Illustrative figures only; not investment advice.",
               italic=True, color=GREY)
        self.t(sh, "B5", "Built by Adham Karim (github.com/adhamkarim4815/lbo-lab)", color=GREY)
        self.t(sh, "B7", "Active case", bold=True)
        self.f(sh, "C7", f"={self.X('Inputs', 'scenario')}", "gen", bold=True, align="left")
        self.t(sh, "B8", "Model status", bold=True)
        self.f(sh, "C8", f"={self.X('Checks', 'master', 'F')}", "gen", bold=True, align="left")
        self.t(sh, "B9", "Sponsor IRR / MOIC", bold=True)
        self.f(sh, "C9", f'=TEXT({self.X("Returns", "irr")},"0.0%")&"  /  "&TEXT({self.X("Returns", "moic")},"0.00")&"x"',
               "gen", bold=True, align="left")
        r = 11
        fm = self.fmt(color=WHITE, bold=True, bg=NAVY)
        for c in "BC":
            ws.write_blank(f"{c}{r}", None, fm)
        self.t(sh, f"B{r}", "Sheet", color=WHITE, bold=True, bg=NAVY)
        self.t(sh, f"C{r}", "Purpose", color=WHITE, bold=True, bg=NAVY)
        guide = [("Summary", "One-page returns summary: S&U, bridge, credit statistics and charts."),
                 ("Inputs", "Every hard-coded assumption, the scenario dropdown (cell D5) and the three operating cases."),
                 ("Transaction", "Purchase price, fees, sources and uses, capitalization and ownership."),
                 ("Operating", "Revenue to EBITDA, taxes, free cash flow before and after debt service."),
                 ("Debt", "Cash waterfall, revolver, term loan (amortization + sweep), PIK-able notes, credit statistics."),
                 ("Returns", "Exit valuation, sponsor IRR and MOIC, option pool, value creation bridge."),
                 ("Sensitivity", "Live-formula IRR and MOIC grids: entry vs exit multiple, exit multiple vs exit year."),
                 ("Checks", "Integrity checks: sources = uses, balances, roll-forwards, bridge, grid reconciliation.")]
        for i, (name, text) in enumerate(guide, start=1):
            ws.write_url(f"B{r + i}", f"internal:'{name}'!A1", self.fmt(color=STEEL), string=name)
            self.t(sh, f"C{r + i}", text)
        r2 = r + len(guide) + 2
        for c in "BC":
            ws.write_blank(f"{c}{r2}", None, fm)
        self.t(sh, f"B{r2}", "Colour code", color=WHITE, bold=True, bg=NAVY)
        ws.write_string(f"B{r2 + 1}", "1,234.5", self.fmt("gen", BLUE, bg=INPUT_BG, box=True))
        self.t(sh, f"C{r2 + 1}", "Blue on cream: hard-coded input. Change these (Inputs sheet only).")
        ws.write_string(f"B{r2 + 2}", "1,234.5", self.fmt("gen", BLACK))
        self.t(sh, f"C{r2 + 2}", "Black: formula or calculation on the same sheet.")
        ws.write_string(f"B{r2 + 3}", "1,234.5", self.fmt("gen", GREEN))
        self.t(sh, f"C{r2 + 3}", "Green: direct link to another sheet.")
        r3 = r2 + 5
        for c in "BC":
            ws.write_blank(f"{c}{r3}", None, fm)
        self.t(sh, f"B{r3}", "Conventions", color=WHITE, bold=True, bg=NAVY)
        notes = ["USD millions unless stated. Year 0 is the closing balance sheet (LTM period); years 1 to 7 are projections.",
                 "Interest accrues on opening balances, so there are no circular references and no iterative calculation.",
                 "Cash waterfall: free cash flow after mandatory amortization, then revolver draw or repayment, then the cash sweep.",
                 "Sweep = sweep % x excess cash flow (period cash flow after mandatory amortization and revolver repayment).",
                 "Bridge: EBITDA growth at the entry multiple; multiple expansion on exit EBITDA; net debt paydown; fees. Parts sum exactly.",
                 "Switch case with the dropdown in Inputs!D5. The Python engine (lbo run) reproduces every number here.",
                 "Halcyon Packaging Co. is fictional. Fees and taxes are simplified; see the repository README for limitations."]
        for i, text in enumerate(notes, start=1):
            self.t(sh, f"B{r3 + i}", f"{i}.", align="right", color=GREY)
            self.t(sh, f"C{r3 + i}", text)
        ws.set_first_sheet()
        ws.activate()


# --------------------------------------------------------------------------- API


def build_workbook(path: str | Path, a: Assumptions, cache: dict | None = None) -> Layout:
    """Write the workbook (formulas only, unless ``cache`` supplies computed values)."""
    return Book(Path(path), a, cache).build()


def read_cache(path: str | Path) -> dict[tuple[str, str], object]:
    wb = openpyxl.load_workbook(path, data_only=True)
    out: dict[tuple[str, str], object] = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[(ws.title, c.coordinate)] = c.value
    return out


def build_model_file(path: str | Path, a: Assumptions, recalculate: bool = True) -> Layout:
    """Build the workbook; if LibreOffice is available also embed cached formula results."""
    path = Path(path)
    layout = build_workbook(path, a)
    if recalculate and recalc.find_soffice():
        with tempfile.TemporaryDirectory() as tmp:
            out = recalc.recalculate(path, Path(tmp) / "out")
            layout = build_workbook(path, a, read_cache(out))
    return layout
