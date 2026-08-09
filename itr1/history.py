"""Prior-year Indian tax parameters used by the Form 10E (relief u/s 89)
arrear-salary computation and the year-on-year comparison view.

Each entry: slab table ``(upper_limit_or_None, rate)`` applied on total income,
standard deduction, section 87A rebate (max, income limit), and cess rate.
Figures follow the Finance Acts of the respective years (general category,
super-senior/senior base adjustments handled by the caller).
"""

from __future__ import annotations

import math

# ---------------------------------------------------------------------------
# (fy_label, slabs_old, std_ded_old, rebate_old, slabs_new, std_new, rebate_new, cess)
#   rebate_* = (max_rebate, income_limit) ; slabs_new None → new regime not offered
# ---------------------------------------------------------------------------

_OLD_SLABS_CLASSIC = ((250000, 0.0), (500000, 0.05), (1000000, 0.20), (None, 0.30))
_NEW_SLABS_2020 = ((250000, 0.0), (500000, 0.05), (750000, 0.10), (1000000, 0.15),
                   (1250000, 0.20), (1500000, 0.25), (None, 0.30))
_NEW_SLABS_2023 = ((300000, 0.0), (600000, 0.05), (900000, 0.10), (1200000, 0.15),
                   (1500000, 0.20), (None, 0.30))
_NEW_SLABS_2024 = ((300000, 0.0), (700000, 0.05), (1000000, 0.10), (1200000, 0.15),
                   (1500000, 0.20), (None, 0.30))

YEARS = {
    "2017-18": dict(old=_OLD_SLABS_CLASSIC, std_old=0,   reb_old=(2500, 350000),
                    new=None, std_new=0, reb_new=None, cess=0.03),
    "2018-19": dict(old=_OLD_SLABS_CLASSIC, std_old=0,   reb_old=(2500, 350000),
                    new=None, std_new=0, reb_new=None, cess=0.04),
    "2019-20": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=None, std_new=0, reb_new=None, cess=0.04),
    "2020-21": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=_NEW_SLABS_2020, std_new=0, reb_new=(12500, 500000), cess=0.04),
    "2021-22": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=_NEW_SLABS_2020, std_new=0, reb_new=(12500, 500000), cess=0.04),
    "2022-23": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=_NEW_SLABS_2020, std_new=0, reb_new=(12500, 500000), cess=0.04),
    "2023-24": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=_NEW_SLABS_2023, std_new=50000, reb_new=(25000, 700000), cess=0.04),
    "2024-25": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=_NEW_SLABS_2024, std_new=75000, reb_new=(25000, 700000), cess=0.04),
    "2025-26": dict(old=_OLD_SLABS_CLASSIC, std_old=50000, reb_old=(12500, 500000),
                    new=((400000, 0.0), (800000, 0.05), (1200000, 0.10), (1600000, 0.15),
                         (2000000, 0.20), (2400000, 0.25), (None, 0.30)),
                    std_new=75000, reb_new=(60000, 1200000), cess=0.04),
}

SENIOR_BASE_OLD = 300000      # resident senior (60+): base exemption in old regime
SUPER_SENIOR_BASE_OLD = 500000  # resident super senior (80+)


def _adjust_for_age(slabs, base):
    """Rebuild an old-regime slab table with a shifted basic exemption."""
    extra = base - 250000
    merged = [((250000 + extra), 0.0), (500000 + extra, 0.05), (1000000 + extra, 0.20),
              (None, 0.30)]
    return tuple(merged)


def _tax_on(income: float, slabs) -> float:
    income = max(0, income)
    tax, prev = 0.0, 0
    for upper, rate in slabs:
        top = income if upper is None else min(income, upper)
        if top > prev:
            tax += (top - prev) * rate
        prev = top if upper is not None else prev
        if upper is None or income <= upper:
            break
    return tax


def tax_for_year(fy: str, total_income: float, *, dob_year: int | None = None,
                 regime: str = "new", salaried: bool = True) -> float:
    """Approximate tax for FY ``fy`` (used only for relief-89 workings)."""
    params = YEARS.get(fy)
    if params is None:
        raise KeyError(f"No tax parameters stored for FY {fy}")
    income = max(0, float(total_income))
    use_new = regime == "new" and params["new"] is not None
    std = (params["std_new"] if use_new else params["std_old"]) if salaried else 0
    taxable = max(0, income - std)
    slabs = params["new"] if use_new else params["old"]
    if not use_new and dob_year is not None:
        # age at the END of that financial year
        y0 = int(fy[:4])
        age = (y0 + 1) - dob_year
        if age >= 80:
            slabs = _adjust_for_age(slabs, SUPER_SENIOR_BASE_OLD)
        elif age >= 60:
            slabs = _adjust_for_age(slabs, SENIOR_BASE_OLD)
    tax = _tax_on(taxable, slabs)
    reb = params["reb_new"] if use_new else params["reb_old"]
    if reb and taxable <= reb[1]:
        tax = max(0.0, tax - reb[0])
    return tax * (1 + params["cess"])


def available_years() -> list[str]:
    return list(YEARS.keys())
