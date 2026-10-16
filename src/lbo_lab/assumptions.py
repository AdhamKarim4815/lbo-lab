"""Transaction assumptions and operating scenarios for the Halcyon Packaging LBO.

All money amounts are USD millions. Halcyon Packaging Co. is a fictional company;
every number here is illustrative.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace

HORIZON = 7
SCENARIO_KEYS = ("base", "management", "downside")


@dataclass(frozen=True)
class Case:
    """One operating scenario: yearly growth and margin paths plus capex and exit multiple."""

    name: str
    revenue_growth: tuple[float, ...]
    ebitda_margin: tuple[float, ...]
    capex_pct: float
    exit_multiple: float


def default_cases() -> dict[str, Case]:
    return {
        "base": Case(
            name="Base",
            revenue_growth=(0.060, 0.055, 0.050, 0.045, 0.040, 0.040, 0.035),
            ebitda_margin=(0.160, 0.164, 0.168, 0.172, 0.175, 0.178, 0.180),
            capex_pct=0.035,
            exit_multiple=9.0,
        ),
        "management": Case(
            name="Management",
            revenue_growth=(0.075, 0.075, 0.065, 0.060, 0.055, 0.050, 0.045),
            ebitda_margin=(0.165, 0.171, 0.177, 0.182, 0.187, 0.191, 0.194),
            capex_pct=0.035,
            exit_multiple=9.5,
        ),
        "downside": Case(
            name="Downside",
            revenue_growth=(0.010, -0.060, 0.000, 0.020, 0.030, 0.030, 0.030),
            ebitda_margin=(0.153, 0.135, 0.140, 0.146, 0.150, 0.153, 0.155),
            capex_pct=0.040,
            exit_multiple=8.0,
        ),
    }


@dataclass(frozen=True)
class Assumptions:
    """Every driver of the model. Defaults describe the Halcyon base transaction."""

    company: str = "Halcyon Packaging Co."
    base_year: int = 2026
    ltm_revenue: float = 625.0
    ltm_ebitda: float = 100.0
    entry_multiple: float = 9.0
    transaction_fee_pct: float = 0.02
    financing_fee_pct: float = 0.025
    fee_amort_years: int = 7
    min_cash: float = 20.0
    rollover_equity: float = 25.0
    option_pool_pct: float = 0.05

    revolver_commitment: float = 60.0
    revolver_draw_at_close: float = 0.0
    term_loan_x: float = 4.0
    sub_notes_x: float = 1.5
    base_rate: float = 0.040
    revolver_spread: float = 0.040
    term_spread: float = 0.0475
    commitment_fee: float = 0.005
    notes_cash_rate: float = 0.11
    notes_pik_rate: float = 0.12
    notes_pik: bool = False
    interest_income_rate: float = 0.02
    sweep_pct: float = 0.75
    mandatory_amort: tuple[float, ...] = (0.025,) * HORIZON

    da_pct: float = 0.03
    nwc_pct: float = 0.12
    tax_rate: float = 0.25

    exit_year: int = 5
    exit_multiple: float | None = None
    interest_method: str = "opening"
    horizon: int = HORIZON
    scenario: str = "base"
    cases: dict[str, Case] = field(default_factory=default_cases)

    @property
    def case(self) -> Case:
        return self.cases[self.scenario]

    @property
    def effective_exit_multiple(self) -> float:
        return self.exit_multiple if self.exit_multiple is not None else self.case.exit_multiple

    def validate(self) -> None:
        if self.scenario not in self.cases:
            raise ValueError(f"unknown scenario {self.scenario!r}; choose from {sorted(self.cases)}")
        if not 1 <= self.horizon <= HORIZON:
            raise ValueError(f"horizon must be between 1 and {HORIZON}")
        if not 1 <= self.exit_year <= self.horizon:
            raise ValueError(f"exit_year must be between 1 and the horizon ({self.horizon})")
        if self.interest_method not in ("opening", "average"):
            raise ValueError("interest_method must be 'opening' or 'average'")
        c = self.case
        for series in (c.revenue_growth, c.ebitda_margin, self.mandatory_amort):
            if len(series) < self.horizon:
                raise ValueError("yearly assumption series are shorter than the horizon")
        if self.revolver_draw_at_close > self.revolver_commitment:
            raise ValueError("revolver draw at close exceeds the commitment")

    def with_scenario(self, scenario: str) -> Assumptions:
        key = resolve_scenario(scenario)
        return replace(self, scenario=key)


def resolve_scenario(name: str) -> str:
    """Map a user-facing name such as 'Management case' or 'base' to a scenario key."""
    key = name.strip().lower().replace(" case", "")
    if key not in SCENARIO_KEYS:
        raise ValueError(f"unknown scenario {name!r}; choose from {', '.join(SCENARIO_KEYS)}")
    return key
