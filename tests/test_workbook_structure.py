"""Structural tests on the generated workbook; these do not need LibreOffice."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from lbo_lab import excel
from lbo_lab.assumptions import Assumptions

ROOT = Path(__file__).resolve().parents[1]
CALC_SHEETS = ["Transaction", "Operating", "Debt", "Returns", "Sensitivity", "Summary", "Checks"]


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    path = tmp_path_factory.mktemp("wb") / "m.xlsx"
    layout = excel.build_workbook(path, Assumptions())
    return layout, openpyxl.load_workbook(path)


def test_sheet_order(built):
    _, wb = built
    assert wb.sheetnames == ["Cover", "Summary", "Inputs", "Transaction", "Operating", "Debt", "Returns",
                             "Sensitivity", "Checks"]


def test_workbook_is_mostly_live_formulas(built):
    _, wb = built
    n = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith("="))
    assert n > 1000


def test_no_hardcoded_numbers_outside_inputs(built):
    """Numeric constants on calculation sheets must be blue inputs (e.g. the sensitivity exit years)."""
    layout, wb = built
    offenders = []
    for name in CALC_SHEETS:
        for row in wb[name].iter_rows():
            for c in row:
                is_number = isinstance(c.value, (int, float)) and not isinstance(c.value, bool)
                if is_number and (c.font.color is None or c.font.color.rgb != "FF0000FF"):
                    offenders.append((name, c.coordinate, c.value))
    assert offenders == []


def test_colour_code(built):
    layout, wb = built
    inputs = wb["Inputs"]
    ltm = inputs[layout.cell("Inputs", "ltm_ebitda")]
    assert ltm.font.color.rgb == "FF0000FF" and not isinstance(ltm.value, str)
    link = wb["Transaction"][layout.cell("Transaction", "ltm_ebitda")]
    assert str(link.value).startswith("=Inputs!") and link.font.color.rgb == "FF008000"
    calc = wb["Transaction"][layout.cell("Transaction", "ev")]
    assert calc.font.color.rgb == "FF000000"


def test_cross_sheet_pure_links_are_green_and_calcs_are_not(built):
    _, wb = built
    import re
    pure = re.compile(r"^=(INDEX\()?[A-Za-z]+!\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?(,\$?[A-Z]+\$?\d+)?\)?$")
    for name in ("Operating", "Debt", "Returns", "Transaction"):
        for row in wb[name].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and c.font.color is not None:
                    is_green = c.font.color.rgb == "FF008000"
                    if pure.match(c.value) and c.font.color.rgb != "FFFFFFFF":
                        assert is_green, (name, c.coordinate, c.value)
                    if not pure.match(c.value):
                        assert not is_green, (name, c.coordinate, c.value)


def test_scenario_dropdown_and_validations(built):
    layout, wb = built
    ws = wb["Inputs"]
    cells = {str(rng) for dv in ws.data_validations.dataValidation for rng in [dv.sqref]}
    assert "D5" in cells
    dv = next(d for d in ws.data_validations.dataValidation if str(d.sqref) == "D5")
    assert dv.type == "list"
    assert ws["D5"].value == "Base"
    assert [ws[f"D{layout.rows['Inputs'][f'scn{i}']}"].value for i in (1, 2, 3)] == ["Base", "Management", "Downside"]


def test_freeze_panes_on_model_sheets(built):
    _, wb = built
    for name in ("Summary", "Inputs", "Transaction", "Operating", "Debt", "Returns", "Sensitivity", "Checks"):
        assert wb[name].freeze_panes, name


def test_no_circular_dependency_in_interest(built):
    """Interest rows must reference opening balances, never closing balances of the same year."""
    layout, wb = built
    ws = wb["Debt"]
    r = layout.rows["Debt"]
    assert ws[f"E{r['int_tl']}"].value == f"=E{r['tl_open']}*$D${r['tl_rate']}"
    assert ws[f"E{r['int_rev']}"].value == f"=E{r['rev_open']}*$D${r['rev_rate']}"
    assert ws[f"E{r['int_notes']}"].value == f"=E{r['notes_open']}*$D${r['notes_rate']}"


def test_native_charts_present(tmp_path):
    import zipfile
    path = tmp_path / "m.xlsx"
    excel.build_workbook(path, Assumptions())
    with zipfile.ZipFile(path) as z:
        charts = [n for n in z.namelist() if n.startswith("xl/charts/chart")]
    assert len(charts) == 3


def test_horizon_must_be_seven(tmp_path):
    with pytest.raises(ValueError):
        excel.build_workbook(tmp_path / "x.xlsx", Assumptions(horizon=5))


def test_committed_example_has_cached_values_matching_engine():
    from lbo_lab.engine import run_model
    path = ROOT / "examples" / "halcyon_lbo.xlsx"
    if not path.exists():
        pytest.skip("example workbook not generated")
    wb = openpyxl.load_workbook(path, data_only=True)
    res = run_model(Assumptions())
    values = [c.value for row in wb["Returns"].iter_rows() for c in row if isinstance(c.value, float)]
    assert any(abs(v - res.returns.irr) < 1e-9 for v in values), "cached IRR not found in the example"
    statuses = [c.value for row in wb["Checks"].iter_rows() for c in row]
    assert "ALL CHECKS PASS" in statuses


def test_cached_build_roundtrip(tmp_path, have_soffice):
    if not have_soffice:
        pytest.skip("LibreOffice not installed")
    from lbo_lab.engine import run_model
    path = tmp_path / "cached.xlsx"
    layout = excel.build_model_file(path, Assumptions())
    wb = openpyxl.load_workbook(path, data_only=True)
    irr = wb["Returns"][layout.cell("Returns", "irr")].value
    assert irr == pytest.approx(run_model(Assumptions()).returns.irr, abs=1e-9)
    formulas = openpyxl.load_workbook(path)["Returns"][layout.cell("Returns", "irr")].value
    assert str(formulas).startswith("=IF(")   # still a live formula
