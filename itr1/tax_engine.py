"""Income-tax computation engine for AY 2026-27 (PY 2025-26), ITR-1 (Sahaj).

Implements:
  * slab-rate tax under the new (default, u/s 115BAC) and old regimes
  * rebate u/s 87A with new-regime marginal relief (CBDT rule 191)
  * surcharge with marginal relief (Income-tax Act s. 115BAC / First Schedule)
  * Health & Education cess @ 4%
  * rounding u/s 288A (income) and 288B (tax)
  * late fee u/s 234F, fee u/s 234-I for revised returns
  * interest u/s 234A / 234B / 234C using actual challan dates

All public functions operate on integer rupees unless stated otherwise.
"""

from __future__ import annotations

import datetime as _dt

from . import constants as C
from .errors import parse_int as _friendly_int


# ---------------------------------------------------------------------------
# Rounding helpers
# ---------------------------------------------------------------------------

def round_off_income(amount: float) -> int:
    """Section 288A - round total income to the nearest multiple of ten
    (fraction of ten of five or more is rounded up)."""
    import math

    return int(math.floor(amount / 10.0 + 0.5)) * 10


def round_off_tax(amount: float) -> int:
    """Section 288B - round tax etc. to the nearest multiple of ten."""
    import math

    return int(math.floor(amount / 10.0 + 0.5)) * 10


def _round_rupee(amount: float) -> int:
    """Round to the nearest whole rupee (traditional half-up)."""
    import math

    return int(math.floor(amount + 0.5))


def floor_hundred(amount: float) -> int:
    """Rule 119A: ignore any fraction of one hundred rupees for interest base."""
    return int(amount // 100) * 100


# ---------------------------------------------------------------------------
# Age / seniority
# ---------------------------------------------------------------------------

def _parse_date(value: str | _dt.date, default: _dt.date | None = None) -> _dt.date | None:
    """Parse an ISO date; tolerant of empty/invalid strings.

    Returns ``default`` (or None) instead of raising when the input is not a
    valid ``YYYY-MM-DD`` date - this keeps live computation working while the
    user is still typing into a partially-filled form.
    """
    if isinstance(value, _dt.date):
        return value
    if not value:
        return default
    try:
        return _dt.date.fromisoformat(str(value).strip())
    except ValueError:
        return default


def is_senior_citizen(dob: str) -> bool:
    """Age 60+ as on 31-03-2026 -> born on or before 01-04-1966 (rule 13).

    A missing/invalid DOB is treated as "below 60" so computational paths
    never crash on incomplete input (schema validation flags the DOB).
    """
    d = _parse_date(dob)
    if d is None:
        return False
    return d <= _parse_date(C.SENIOR_CITIZEN_DOB_ON_OR_BEFORE)


def is_super_senior_citizen(dob: str) -> bool:
    """Age 80+ as on 31-03-2026 -> born on or before 01-04-1946."""
    d = _parse_date(dob)
    if d is None:
        return False
    return d <= _parse_date(C.SUPER_SENIOR_DOB_ON_OR_BEFORE)


# ---------------------------------------------------------------------------
# Slab tax
# ---------------------------------------------------------------------------

def slab_tax(income: int, slabs: list[tuple[int, float]]) -> float:
    """Progressive slab tax on ``income`` using (lower_limit, rate) pairs."""
    if income <= 0:
        return 0.0
    tax = 0.0
    for idx, (lower, rate) in enumerate(slabs):
        upper = slabs[idx + 1][0] if idx + 1 < len(slabs) else None
        if income <= lower:
            break
        taxable_in_slab = income - lower if upper is None else min(income, upper) - lower
        tax += taxable_in_slab * rate
    return tax


def normal_slab_tax(income: int, opt_out_new_regime: bool, dob: str,
                    residential_status: str = "RES") -> float:
    """Tax at slab rates for the chosen regime and age category.

    Age-based (senior / super-senior) old-regime slabs apply only to
    residents — ``RES`` or ``NOR`` (RNOR is resident u/s 6).  A
    non-resident senior citizen gets the general slab, exactly as the
    official utilities' TaxCalc sheets do::

        IF(AND(Pan="P",taxPayerStatus="I",bacValue=2,age>59,age<80,
               OR(resStatus="RES",resStatus="NOR")), 1, ... 3)

    (ITR-2 utility 'Tax Calculated'!B26 / C34; ITR-4 TaxCalc identical.)
    """
    senior_ok = str(residential_status or "RES").upper() != "NRI"
    if opt_out_new_regime:
        if senior_ok and is_super_senior_citizen(dob):
            slabs = C.OLD_REGIME_SLABS_SUPER_SENIOR
        elif senior_ok and is_senior_citizen(dob):
            slabs = C.OLD_REGIME_SLABS_SENIOR
        else:
            slabs = C.OLD_REGIME_SLABS_GENERAL
    else:
        slabs = C.NEW_REGIME_SLABS
    return slab_tax(income, slabs)


# ---------------------------------------------------------------------------
# Surcharge with marginal relief
# ---------------------------------------------------------------------------

# Tax u/s 115BBE (unexplained income) carries a flat 25% surcharge,
# computed separately from the slab surcharge (same split as the official
# utilities' TaxCalc sheets).
SURCHARGE_115BBE_FLAT = 0.25

def _surcharge_rate(total_income: int, opt_out_new_regime: bool) -> float:
    rate = 0.0
    for threshold, r in C.SURCHARGE:
        if total_income > threshold:
            rate = r
    if not opt_out_new_regime:
        rate = min(rate, C.SURCHARGE_MAX_RATE_NEW_REGIME)
    return rate


def compute_surcharge(base_tax: float, total_income: int, ltcg112a: int,
                      opt_out_new_regime: bool, dob: str,
                      residential_status: str = "RES") -> float:
    """Surcharge on ``base_tax`` honouring marginal relief.

    ``base_tax`` is the aggregate of slab tax on normal income plus the
    special-rate tax on LTCG u/s 112A (zero within ITR-1 limits).
    """
    rate = _surcharge_rate(total_income, opt_out_new_regime)
    if rate == 0.0 or base_tax <= 0:
        return 0.0

    surcharge = base_tax * rate

    # Marginal relief: tax + surcharge must not exceed tax on the threshold
    # income plus the income in excess of the threshold.
    applicable_threshold = 0
    for threshold, r in C.SURCHARGE:
        if total_income > threshold:
            applicable_threshold = threshold
    threshold_normal = max(0, applicable_threshold - min(ltcg112a, applicable_threshold))
    tax_at_threshold = normal_slab_tax(threshold_normal, opt_out_new_regime, dob,
                                       residential_status)
    ltcg_at_threshold = applicable_threshold - threshold_normal
    tax_at_threshold += ltcg_tax_112a(ltcg_at_threshold, threshold_normal)
    # surcharge applicable at the threshold income
    threshold_rate = 0.0
    for threshold, r in C.SURCHARGE:
        if applicable_threshold > threshold:
            threshold_rate = r
    if not opt_out_new_regime:
        threshold_rate = min(threshold_rate, C.SURCHARGE_MAX_RATE_NEW_REGIME)
    max_payable = tax_at_threshold * (1 + threshold_rate) + (total_income - applicable_threshold)
    if base_tax + surcharge > max_payable:
        surcharge = max(0.0, max_payable - base_tax)
    return surcharge


# ---------------------------------------------------------------------------
# LTCG u/s 112A
# ---------------------------------------------------------------------------

def ltcg_tax_112a(ltcg: int, other_income: int = 0) -> float:
    """Tax on LTCG u/s 112A at 12.5% above the Rs. 1.25 lakh exemption.

    Within ITR-1 the reported LTCG 112A cannot exceed Rs. 1,25,000 (schema
    cap / rule 217), so this always evaluates to zero here; implemented for
    completeness and unit-tested at the boundary.
    """
    if ltcg <= 0:
        return 0.0
    exemption = C.MAX_LTCG_112A
    # unexhausted basic exemption may shield the gain further
    taxable = max(0, ltcg - exemption)
    return taxable * C.LTCG_112A_RATE


# ---------------------------------------------------------------------------
# Rebate u/s 87A
# ---------------------------------------------------------------------------

def compute_rebate_87a(total_tax_payable: float, total_income: int,
                       income_excl_ltcg: int, opt_out_new_regime: bool,
                       residential_status: str = "RES",
                       status: str = "I") -> float:
    """Rebate u/s 87A including new-regime marginal relief (rule 191).

    Only a resident individual may claim it — CBDT portal validation
    rules 533/534 ("Rebate u/s 87A is not allowed to Non-resident / HUF")
    and the official utility formula ``Rebate87Aformula_new``::

        IF(MID(Status,1,1)<>"H",
           IF(MID(ResidentialStatus1,1,3)<>"NRI", <rebate>, 0), 0)
    """
    if str(residential_status or "RES").upper() == "NRI":
        return 0.0
    if str(status or "I").upper() != "I":
        return 0.0
    if total_tax_payable <= 0:
        return 0.0
    if opt_out_new_regime:
        # rule 23 / 192: resident with total income (incl. LTCG) up to Rs. 5L
        if total_income <= C.REBATE_87A_OLD_INCOME_LIMIT:
            return min(total_tax_payable, C.REBATE_87A_OLD_MAX)
        return 0.0
    # New regime: rule 191 uses total income excluding LTCG 112A
    if income_excl_ltcg <= C.REBATE_87A_NEW_INCOME_LIMIT:
        return min(total_tax_payable, C.REBATE_87A_NEW_MAX)
    if income_excl_ltcg <= C.REBATE_87A_NEW_MR_CEILING:
        # marginal relief: tax is capped at income in excess of Rs. 12 lakh
        marginal_tax = max(0, income_excl_ltcg - C.REBATE_87A_NEW_INCOME_LIMIT)
        rebate = total_tax_payable - marginal_tax
        return max(0.0, min(rebate, C.REBATE_87A_NEW_MAX))
    return 0.0


# ---------------------------------------------------------------------------
# Fee u/s 234F and fee u/s 234-I
# ---------------------------------------------------------------------------

def fee_234f(total_income: int, filing_date: str, return_section: int) -> int:
    """Late filing fee u/s 234F.

    Rs. 1,000 where total income <= Rs. 5 lakh else Rs. 5,000, levied when the
    return is filed after the due date (31-07-2026) or filed u/s 139(4)/139(5)
    after the due date.
    """
    late = _parse_date(filing_date, _dt.date.today()) > _parse_date(C.DUE_DATE_FILING)
    if not late and return_section not in (12,):
        return 0
    if not late:
        return 0
    return C.LATE_FEE_234F_SMALL if total_income <= 500_000 else C.LATE_FEE_234F_FULL


def fee_234i(total_income: int, filing_date: str, return_section: int) -> int:
    """Fee u/s 234-I for furnishing a revised return after 31-12-2026
    (CBDT rules 324 and 328)."""
    if return_section != 17:
        return 0
    if _parse_date(filing_date, _dt.date.today()) <= _parse_date(C.REVISED_RETURN_FEE_CUTOFF):
        return 0
    return C.FEE_234I_SMALL if total_income <= 500_000 else C.FEE_234I_FULL


# ---------------------------------------------------------------------------
# Interest u/s 234A / 234B / 234C
# ---------------------------------------------------------------------------

def _months_between(start: _dt.date, end: _dt.date) -> int:
    """Number of months or part-months from ``start`` (exclusive) to ``end``."""
    if end <= start:
        return 0
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day > start.day:
        months += 1
    return months


def interest_234a(net_tax_liability: float, taxes_before_due: float,
                  filing_date: str, return_section: int) -> int:
    """Interest u/s 234A: 1% per month (or part) on the unpaid tax from the
    day after the due date up to the date of filing."""
    return _interest_234a(net_tax_liability, taxes_before_due, filing_date,
                          return_section)["amount"]


def _interest_234a(net_tax_liability: float, taxes_before_due: float,
                   filing_date: str, return_section: int) -> dict:
    due = _parse_date(C.DUE_DATE_FILING)
    filed = _parse_date(filing_date, _dt.date.today())
    on_time = filed <= due and return_section == 11
    base = 0 if on_time else floor_hundred(max(0.0, net_tax_liability - taxes_before_due))
    months = 0 if (on_time or base <= 0) else _months_between(due, filed)
    return {
        "base": base,
        "months": months,
        "rate": C.INTEREST_234_RATE,
        "amount": _round_rupee(base * C.INTEREST_234_RATE * months),
        "note": ("No interest — return filed within the due date." if on_time else
                 f"1% p.m. on unpaid Rs. {base:,} for {months} month(s) after the due date."),
    }


def _challans_py(challans: list[dict]) -> list[tuple[_dt.date, int]]:
    """Advance-tax portion of challans (deposited within the PY)."""
    out = []
    for ch in challans or []:
        d = _parse_date(ch.get("date", "1970-01-01"), _dt.date(1970, 1, 1))
        amt = _friendly_int(ch.get("amount", 0), default=0,
                            field="Advance/Self-assessment tax challan amount",
                            path="tax_paid.challans")
        if _dt.date(2025, 4, 1) <= d <= _dt.date(2026, 3, 31):
            out.append((d, amt))
    return sorted(out)


def _challans_after_py(challans: list[dict]) -> int:
    total = 0
    for ch in challans or []:
        d = _parse_date(ch.get("date", "1970-01-01"), _dt.date(1970, 1, 1))
        if d > _dt.date(2026, 3, 31):
            total += _friendly_int(ch.get("amount", 0), default=0,
                                   field="Advance/Self-assessment tax challan amount",
                                   path="tax_paid.challans")
    return total


def interest_234b(assessed_tax: float, advance_tax_paid: float,
                  filing_date: str) -> int:
    """Interest u/s 234B: advance tax below 90% of assessed tax.

    1% per month from 1 April of the AY to the month of payment/filing on the
    shortfall. Simplified: months counted from 01-04-2026 to the filing date.
    """
    return _interest_234b(assessed_tax, advance_tax_paid, filing_date)["amount"]


def _interest_234b(assessed_tax: float, advance_tax_paid: float,
                   filing_date: str) -> dict:
    assessed = floor_hundred(max(0.0, assessed_tax))
    shortfall = 0
    months = 0
    if assessed > 0 and advance_tax_paid < 0.9 * assessed:
        shortfall = floor_hundred(assessed - advance_tax_paid)
        months = _months_between(_dt.date(2026, 3, 31),
                                 _parse_date(filing_date, _dt.date.today()))
        months = max(months, 1)
    note = ("No interest — advance tax covers at least 90% of assessed tax "
            "(or no tax is payable)." if shortfall == 0 else
            f"1% p.m. on shortfall Rs. {shortfall:,} for {months} month(s).")
    return {"assessed_tax": assessed, "advance_tax": int(advance_tax_paid),
            "shortfall": shortfall, "months": months, "rate": C.INTEREST_234_RATE,
            "amount": _round_rupee(shortfall * C.INTEREST_234_RATE * months),
            "note": note}


def interest_234c(assessed_tax: float, tds_tcs: float,
                  challans: list[dict]) -> int:
    """Interest u/s 234C for deferment of advance-tax instalments.

    Uses the actual dates of advance-tax challans. Tax due for each
    instalment is 15% / 45% / 75% / 100% of the tax due on returned income
    (assessed tax less TDS/TCS already standing at that point is simplified
    to assessed tax net of total TDS/TCS).
    """
    return _interest_234c(assessed_tax, tds_tcs, challans)["amount"]


def _interest_234c(assessed_tax: float, tds_tcs: float,
                   challans: list[dict]) -> dict:
    tax_due = max(0.0, assessed_tax - tds_tcs)
    instalments = []
    if tax_due < 10_000:  # section 234B threshold; no advance tax payable
        return {"tax_due": int(tax_due), "instalments": [], "amount": 0,
                "note": "No advance tax was required (tax due on returned income "
                        "is below Rs. 10,000)."}
    paid = _challans_py(challans)
    interest = 0.0
    cumulative = 0
    for idx, (deadline, pct) in enumerate(C.ADVANCE_TAX_INSTALMENTS):
        dl = _parse_date(deadline)
        cumulative += sum(a for d, a in paid if d <= dl)
        required = floor_hundred(tax_due * pct)
        shortfall = floor_hundred(max(0, required - cumulative))
        months = 3 if pct < 1.0 else 1
        amt = _round_rupee(shortfall * C.INTEREST_234_RATE * months)
        instalments.append({
            "due": deadline, "required_pct": pct, "required": required,
            "paid": cumulative, "shortfall": shortfall, "months": months,
            "amount": amt,
        })
        interest += shortfall * C.INTEREST_234_RATE * months
    return {"tax_due": int(tax_due), "instalments": instalments,
            "amount": _round_rupee(interest),
            "note": "1% p.m. on the shortfall of each advance-tax instalment."}


# ---------------------------------------------------------------------------
# Full tax computation for a return
# ---------------------------------------------------------------------------

def compute_tax(*, total_income: int, ltcg112a: int, opt_out_new_regime: bool,
                dob: str, section89: int = 0, filing_date: str,
                return_section: int = 11, tds: float = 0.0, tcs: float = 0.0,
                advance_tax_challans: list[dict] | None = None,
                interest_override: dict | None = None,
                residential_status: str = "RES",
                status: str = "I") -> dict:
    """Compute the complete tax liability for a return (ITR-1 profile).

    ``total_income`` is the rounded total income (including LTCG 112A,
    schema-capped at Rs. 1.25 lakh).  Returns a mapping with every field
    required by ``ITR1_TaxComputation``, ``TaxPaid`` and ``Refund`` plus
    interest detail.  Thin wrapper over :func:`compute_tax_v2`.
    """
    return compute_tax_v2(
        total_income=total_income,
        special={"ltcg_112a": min(max(0, int(ltcg112a)), C.MAX_LTCG_112A)},
        opt_out_new_regime=opt_out_new_regime, dob=dob, section89=section89,
        filing_date=filing_date, return_section=return_section, tds=tds, tcs=tcs,
        advance_tax_challans=advance_tax_challans,
        interest_override=interest_override,
        residential_status=residential_status, status=status)


def compute_tax_v2(*, total_income: int, special: dict | None = None,
                   opt_out_new_regime: bool,
                   dob: str, section89: int = 0, filing_date: str,
                   return_section: int = 11, tds: float = 0.0, tcs: float = 0.0,
                   advance_tax_challans: list[dict] | None = None,
                   interest_override: dict | None = None,
                   residential_status: str = "RES",
                   status: str = "I") -> dict:
    """Complete tax liability with special-rate capital-gain buckets (ITR-2/4).

    ``special`` keys (amounts already included in ``total_income``):
      stcg_111a  — STCG on equity/equity-MF with STT (20%),
      ltcg_112a  — LTCG u/s 112A (12.5% above Rs. 1.25 lakh),
      ltcg_125   — other LTCG w/o indexation (12.5%),
      ltcg_idx20 — LTCG at 20% with indexation (pre-23-07-2024 option),
                   os_special — list of OS special-rate rows (Schedule SI, see
                   itr1.special_rates): amounts included in total income;
                   tax u/s 115BBE additionally carries a flat 25%
                   surcharge computed separately (utility-TaxCalc parity).
    ``residential_status`` (``RES``/``NOR``/``NRI``) and ``status``
    (``I`` individual / ``H`` HUF …) mirror the official utilities' TaxCalc:
    age-based old-regime slabs apply only to residents (``RES``/``NOR``),
    and rebate u/s 87A is denied to non-residents and non-individuals
    (portal validation rules 533/534).
    """
    residential_status = str(residential_status or "RES").upper()
    status = str(status or "I").upper()
    special = special or {}
    ltcg112a = max(0, int(special.get("ltcg_112a", 0)))
    s111a = max(0, int(special.get("stcg_111a", 0)))
    l125 = max(0, int(special.get("ltcg_125", 0)))
    lidx = max(0, int(special.get("ltcg_idx20", 0)))
    cg_total = ltcg112a + s111a + l125 + lidx
    # OS-family special-rate incomes (Schedule SI): rows from
    # itr1.special_rates.os_special_rows -> {code, amount, rate, tax}
    os_rows = list(special.get("os_special") or [])
    os_total = sum(int(r.get("amount", 0)) for r in os_rows)
    os_tax = sum(int(r.get("tax", 0)) for r in os_rows)
    bbe_tax = sum(int(r.get("tax", 0)) for r in os_rows
                  if r.get("code") == "5BBE")
    bbe_amount = sum(int(r.get("amount", 0)) for r in os_rows
                     if r.get("code") == "5BBE")
    special_total = cg_total + os_total
    total_income = max(0, int(total_income))
    income_excl_ltcg = max(0, total_income - special_total)

    slab = normal_slab_tax(income_excl_ltcg, opt_out_new_regime, dob,
                           residential_status)
    ltcg_tax = ltcg_tax_112a(ltcg112a, income_excl_ltcg)
    tax_111a = s111a * C.STCG_111A_RATE
    tax_125 = l125 * C.LTCG_GENERAL_125_RATE
    tax_idx = lidx * C.LTCG_INDEXED_20_RATE
    special_tax = ltcg_tax + tax_111a + tax_125 + tax_idx + os_tax
    base_tax = slab + special_tax

    # Surcharge: tax u/s 115BBE carries a FLAT 25% surcharge (computed
    # separately, exactly as the utility's TaxCalc does: "Add surcharge on
    # 115BBE" / "Surcharge on other than 115BBE"); the remainder follows the
    # normal slabs-with-marginal-relief path.
    if bbe_tax:
        surcharge_115bbe = bbe_tax * SURCHARGE_115BBE_FLAT
        surcharge_rest = compute_surcharge(base_tax - bbe_tax, total_income,
                                           special_total - bbe_amount,
                                           opt_out_new_regime, dob,
                                           residential_status)
        surcharge = surcharge_rest + surcharge_115bbe
    else:
        surcharge_115bbe = 0.0
        surcharge = compute_surcharge(base_tax, total_income, special_total,
                                      opt_out_new_regime, dob,
                                      residential_status)
    total_tax_payable = base_tax + surcharge

    rebate = compute_rebate_87a(total_tax_payable, total_income,
                                income_excl_ltcg, opt_out_new_regime,
                                residential_status, status)
    tax_after_rebate = total_tax_payable - rebate

    cess = tax_after_rebate * C.HEALTH_EDUCATION_CESS
    gross_tax_liability = tax_after_rebate + cess

    relief89 = max(0, _friendly_int(section89, default=0,
                                    field="Relief u/s 89", path="relief_89"))
    net_tax_liability = max(0.0, gross_tax_liability - relief89)

    # rounding u/s 288B at the liability level
    net_tax_liability_r = float(round_off_tax(net_tax_liability))

    challans = advance_tax_challans or []
    advance_tax = sum(a for _, a in _challans_py(challans))
    self_assessment = _challans_after_py(challans)

    interest_ = {
        "IntrstPayUs234A": interest_234a(net_tax_liability_r, tds + tcs + advance_tax,
                                         filing_date, return_section),
        "IntrstPayUs234B": interest_234b(net_tax_liability_r - tds - tcs,
                                         advance_tax, filing_date),
        "IntrstPayUs234C": interest_234c(net_tax_liability_r, tds + tcs, challans),
        "LateFilingFee234F": fee_234f(total_income, filing_date, return_section),
        "FeeFurnish234I": fee_234i(total_income, filing_date, return_section),
    }
    detail = {
        "IntrstPayUs234A": _interest_234a(net_tax_liability_r, tds + tcs + advance_tax,
                                          filing_date, return_section),
        "IntrstPayUs234B": _interest_234b(net_tax_liability_r - tds - tcs,
                                          advance_tax, filing_date),
        "IntrstPayUs234C": _interest_234c(net_tax_liability_r, tds + tcs, challans),
        "LateFilingFee234F": {
            "amount": interest_["LateFilingFee234F"],
            "note": (f"Late filing fee Rs. {interest_['LateFilingFee234F']:,} "
                     f"(filed after the due date; total income "
                     f"{'up to' if total_income <= 500_000 else 'above'} Rs. 5 lakh)."
                     if interest_["LateFilingFee234F"] else
                     "No fee — return filed on or before the due date."),
        },
        "FeeFurnish234I": {
            "amount": interest_["FeeFurnish234I"],
            "note": ("Fee for furnishing a revised return after 31-12-2026."
                     if interest_["FeeFurnish234I"] else
                     "No fee u/s 234-I."),
        },
    }
    overridden = set()
    if interest_override:
        for key, value in interest_override.items():
            if key in interest_ and value not in (None, ""):
                nice = {"IntrstPayUs234A": "Interest u/s 234A",
                        "IntrstPayUs234B": "Interest u/s 234B",
                        "IntrstPayUs234C": "Interest u/s 234C",
                        "LateFilingFee234F": "Late filing fee u/s 234F",
                        "FeeFurnish234I": "Fee u/s 234-I"}.get(key, key)
                interest_[key] = max(0, _friendly_int(
                    value, default=0, field=f"Manual override: {nice}",
                    path=f"interest_override.{key}"))
                if key in detail:
                    detail[key]["amount"] = interest_[key]
                    detail[key]["note"] = "Manually overridden."
                    overridden.add(key)

    total_interest_fee = sum(interest_.values())
    tot_tax_plus_interest = round_off_tax(net_tax_liability + total_interest_fee)

    total_taxes_paid = int(round(tds + tcs + advance_tax + self_assessment))
    balance_payable = max(0, tot_tax_plus_interest - total_taxes_paid)
    refund_due = max(0, total_taxes_paid - tot_tax_plus_interest)

    return {
        "TotalTaxPayable": _round_rupee(total_tax_payable),
        "SurchargeComputed": _round_rupee(surcharge),
        "Rebate87A": _round_rupee(rebate),
        "TaxPayableOnRebate": _round_rupee(tax_after_rebate),
        "EducationCess": _round_rupee(cess),
        "GrossTaxLiability": _round_rupee(gross_tax_liability),
        "Section89": relief89,
        "NetTaxLiability": int(net_tax_liability_r),
        "IntrstPay": interest_,
        "InterestDetail": detail,
        "InterestOverridden": sorted(overridden),
        "TotalIntrstPay": int(total_interest_fee),
        "TotTaxPlusIntrstPay": int(tot_tax_plus_interest),
        "TaxesPaid": {
            "AdvanceTax": int(advance_tax),
            "TDS": int(round(tds)),
            "TCS": int(round(tcs)),
            "SelfAssessmentTax": int(self_assessment),
            "TotalTaxesPaid": total_taxes_paid,
        },
        "BalTaxPayable": int(balance_payable),
        "RefundDue": int(refund_due),
        # explanation material for the computation report
        "Working": {
            "income_excl_ltcg": income_excl_ltcg,
            "slab_tax": _round_rupee(slab),
            "ltcg_tax": _round_rupee(ltcg_tax),
            "tax_111a": _round_rupee(tax_111a),
            "tax_ltcg_125": _round_rupee(tax_125),
            "tax_ltcg_idx20": _round_rupee(tax_idx),
            "special": {"stcg_111a": s111a, "ltcg_112a": ltcg112a,
                        "ltcg_125": l125, "ltcg_idx20": lidx},
            "os_special": [
                {"code": r.get("code"), "label": r.get("label"),
                 "amount": int(r.get("amount", 0)), "rate": r.get("rate"),
                 "tax": _round_rupee(r.get("tax", 0))} for r in os_rows],
            "tax_os_special": _round_rupee(os_tax),
            "surcharge_115bbe": _round_rupee(surcharge_115bbe),
            "base_tax": _round_rupee(base_tax),
            "residential_status": residential_status,
            "status": status,
            "rebate_note": rebate_note(income_excl_ltcg, total_income,
                                       opt_out_new_regime,
                                       residential_status, status),
        },
    }


def rebate_note(income_excl_ltcg: int, total_income: int, opt_out: bool,
                residential_status: str = "RES", status: str = "I") -> str:
    """Human-readable explanation of the 87A rebate applied."""
    if str(residential_status or "").upper() == "NRI":
        return ("No rebate u/s 87A — not available to a non-resident "
                "(portal validation rule 533).")
    if str(status or "I").upper() != "I":
        return ("No rebate u/s 87A — available only to an individual "
                "(portal validation rule 534).")
    if opt_out:
        if total_income <= C.REBATE_87A_OLD_INCOME_LIMIT:
            return ("Total income is within Rs. 5,00,000, so rebate u/s 87A up to "
                    "Rs. 12,500 applies (old regime).")
        return "Total income exceeds Rs. 5,00,000, so no rebate u/s 87A (old regime)."
    if income_excl_ltcg <= C.REBATE_87A_NEW_INCOME_LIMIT:
        return ("Total income (excl. LTCG 112A) is within Rs. 12,00,000, so rebate "
                "u/s 87A up to Rs. 60,000 applies (new regime).")
    if income_excl_ltcg <= C.REBATE_87A_NEW_MR_CEILING:
        return ("Marginal relief applied: tax is capped at the income exceeding "
                "Rs. 12,00,000 (new regime, rule 191).")
    return "Total income (excl. LTCG 112A) exceeds Rs. 12,70,590, so no rebate u/s 87A."


def slab_breakup(income: int, opt_out_new_regime: bool, dob: str) -> list[dict]:
    """Human-readable slab-wise breakup (for the UI / CLI report)."""
    if opt_out_new_regime:
        if is_super_senior_citizen(dob):
            slabs = C.OLD_REGIME_SLABS_SUPER_SENIOR
        elif is_senior_citizen(dob):
            slabs = C.OLD_REGIME_SLABS_SENIOR
        else:
            slabs = C.OLD_REGIME_SLABS_GENERAL
    else:
        slabs = C.NEW_REGIME_SLABS
    rows = []
    for idx, (lower, rate) in enumerate(slabs):
        upper = slabs[idx + 1][0] if idx + 1 < len(slabs) else None
        if income <= lower:
            continue
        taxable = income - lower if upper is None else min(income, upper) - lower
        rows.append({
            "slab": f"{lower + 1:,} - {upper:,}" if upper else f"Above {lower:,}",
            "rate": rate,
            "taxable": taxable,
            "tax": taxable * rate,
        })
    return rows
