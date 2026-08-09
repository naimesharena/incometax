"""Special-rate income (Schedule SI) — AY 2026-27.

``SEC_CODES`` mirrors, verbatim, the ``SecCode`` enumeration documented in
the official ITR-2 schema (``ScheduleSI.SplCodeRateTax``).  Rates and the
surcharge nuance are cross-checked against the official utilities' hidden
``Tax Calculated`` sheets (see ``extracted/xlsm/*/sheets/`` — the macros
annotate the same figures: 111A '15% & 20%' across the 23-07-2024 change,
112/112A/PTI 12.5%, 115A/115AC 'normal 20%', 'normal 30%', and a separately
computed flat 25% surcharge on tax u/s 115BBE).

Wiring status (kept honest — anything not wired end-to-end is validated as
an unsupported code rather than silently producing a wrong document):

* CG-family codes 1A / 2A / 21(12.5%) / 21(20%-indexed) are computed by the
  capital-gains engine (:mod:`itr1.capgains`) end-to-end.
* OS-family codes 5BB / 5BBJ / 5BBE are wired end-to-end here
  (input → Schedule OS → Part B-TI ``IncChargblSplRateOS`` → Schedule SI →
  tax engine incl. the 115BBE flat-25% surcharge).  5BBA (non-resident
  sportsmen) has no dedicated Schedule-OS node and is therefore reported
  as unsupported rather than approximated.
* Every other code raises a helpful error pointing at the remainder
  (115A-family, FII/FPI 115AD, 115BBF/BBG/BBH-VDA, 115E, DTAA/PTI rows),
  which is documented scope — see ``UNSUPPORTED_NOTE``.
"""
from __future__ import annotations

from .errors import DataError, parse_float

# schema SecCode -> (label per schema, statutory rate % (None = DTAA/situational))
SEC_CODES: dict[str, tuple[str, float | None]] = {
    "1": ("Tax on accumulated balance of recognised PF", None),
    "1A": ("STCG u/s 111A on shares/units (STT paid)", 20.0),
    "21": ("LTCG u/s 112", None),
    "22": ("LTCG on listed securities/units (112(1) proviso)", None),
    "21ciii": ("LTCG on unlisted securities in case of non-residents "
               "(112(1)(c)(iii))", None),
    "2A": ("LTCG u/s 112A on equity (STT paid)", 12.5),
    "5A1ai": ("115A(1)(a)(i) dividends/interest/units in foreign currency", 20.0),
    "5A1aA": ("115A(1)(a)(A) dividend from an IFSC unit - NR", 10.0),
    "5A1aii": ("115A(1)(a)(ii) interest from Govt/Indian concerns (FC)", 20.0),
    "5A1aiia": ("115A(1)(a)(iia) interest from Infrastructure Debt Fund", 5.0),
    "5A1aiiaa": ("115A(1)(a)(iiaa) income referred in 194LC(1)", 5.0),
    "5A1aiiaaP": ("115A(1)(a)(iiaa) proviso to 194LC(1)", 4.0),
    "5A1aiiaa2P": ("115A(1)(a)(iiaa) second proviso to 194LC(1)", 4.0),
    "5A1aiiab": ("115A(1)(a)(iiab) interest u/s 194LD", 5.0),
    "5A1aiiac": ("115A(1)(a)(iiac) interest u/s 194LBA", 5.0),
    "5A1aiii": ("115A(1)(a)(iii) units of UTI in foreign currency", 20.0),
    "5A1bA": ("115A(1)(b)(A)/(B) royalty & technical services", 10.0),
    "5AC1ab": ("115AC(1)(a) interest on bonds in foreign currency - NR", 10.0),
    "5AC1abD": ("115AC(1)(b) dividend on GDRs in foreign currency - NR", 10.0),
    "5AC1c": ("115AC(1)(c) LTCG on bonds/GDR in foreign currency - NR", 12.5),
    "5ACA1a": ("115ACA(1)(a) GDR income in foreign currency - resident", 10.0),
    "5ACA1b": ("115ACA(1)(b) LTCG on GDR in foreign currency - resident", 12.5),
    "5AD1i": ("115AD(1)(i) FII income other than dividend", 20.0),
    "5AD1iDiv": ("115AD(1)(i) FII dividend income", 20.0),
    "5AD1iP": ("115AD(1)(i) FII bonds/Govt securities income u/s 194LD", 5.0),
    "5ADii": ("115AD(1)(ii) FII STCG (other than 111A)", 30.0),
    "5AD1biip": ("115AD(1)(b)(ii) FII STCG u/s 111A", 20.0),
    "5ADiii": ("115AD(1)(iii) FII LTCG", 12.5),
    "5ADiiiP": ("115AD(1)(b)(iii) proviso FII LTCG u/s 112A (STT paid)", 12.5),
    "5BB": ("115BB winnings from lotteries/crossword/races/card games/"
            "gambling/betting", 30.0),
    "5BBJ": ("115BBJ winnings from online games", 30.0),
    "5BBA": ("115BBA non-resident sportsmen/sports associations", 20.0),
    "5BBE": ("115BBE unexplained income u/s 68/69/69A/69B/69C/69D", 60.0),
    "5BBF": ("115BBF income from patent", 10.0),
    "5BBG": ("115BBG income from transfer of carbon credits", 10.0),
    "5BBH": ("115BBH virtual digital assets (VDA)", 30.0),
    "5Ea": ("115E(a) investment income - NR", 20.0),
    "5Eb": ("115E(b) LTCG on foreign exchange assets - NRI", 12.5),
    "DTAASTCG": ("STCG chargeable at DTAA rate", None),
    "DTAALTCG": ("LTCG chargeable at DTAA rate", None),
    "DTAAOS": ("other-source income chargeable at DTAA rate", None),
    # Pass-through income (PTI) variants — same rates as the base codes.
    "PTI_STCG20P": ("PTI in the nature of STCG @20% u/s 111A", 20.0),
    "PTI_STCG30P": ("PTI in the nature of STCG @30%", 30.0),
    "PTI_LTCG12_5P112A": ("PTI in the nature of LTCG @12.5% u/s 112A", 12.5),
    "PTI_LTCG12_5P": ("PTI in the nature of LTCG @12.5%", 12.5),
    "PTI_5A1ai": ("PTI-115A(1)(a)(i) dividends/interest/units in FC", 20.0),
    "PTI_5A1aA": ("PTI-115A(1)(a)(A) dividend from IFSC unit - NR", 10.0),
    "PTI_5A1aii": ("PTI-115A(1)(a)(ii) interest from Govt/Indian concerns", 20.0),
    "PTI_5A1aiia": ("PTI-115A(1)(a)(iia) interest from Infrastructure Debt Fund", 5.0),
    "PTI_5A1aiiaa": ("PTI-115A(1)(a)(iiaa) income u/s 194LC(1)", 5.0),
    "PTI_5A1aiiaaP": ("PTI-115A(1)(a)(iiaa) proviso to 194LC(1)", 4.0),
    "PTI_5A1aiiaa2P": ("PTI-115A(1)(a)(iiaa) second proviso to 194LC(1)", 4.0),
    "PTI_5A1aiiab": ("PTI-115A(1)(a)(iiab) interest u/s 194LD", 5.0),
    "PTI_5A1aiiac": ("PTI-115A(1)(a)(iiac) interest u/s 194LBA", 5.0),
    "PTI_5A1aiii": ("PTI-115A(1)(a)(iii) units of UTI in FC", 20.0),
    "PTI_5A1bA": ("PTI-115A(1)(b)(A)/(B) royalty & technical services", 10.0),
    "PTI_5AC1ab": ("PTI-115AC(1)(a) interest on bonds in FC - NR", 10.0),
    "PTI_5AC1abD": ("PTI-115AC(1)(b) dividend on GDRs in FC - NR", 10.0),
    "PTI_5ACA1a": ("PTI-115ACA(1)(a) GDR income in FC - resident", 10.0),
    "PTI_5AD1i": ("PTI-115AD(1)(i) FII income other than dividend", 20.0),
    "PTI_5AD1iDiv": ("PTI-115AD(1)(i) FII dividend income", 20.0),
    "PTI_5AD1iP": ("PTI-115AD(1)(i) FII bonds/Govt securities u/s 194LD", 5.0),
    "PTI_5BBA": ("PTI-115BBA non-resident sportsmen/associations", 20.0),
    "PTI_5BBG": ("PTI-115BBG transfer of carbon credits", 10.0),
    "PTI_5Ea": ("PTI-115E(a) investment income", 20.0),
}

# OS-family keys wired end-to-end: friendly input key ->
# (schema SecCode, Schedule-OS top node, quarterly daterange node or None).
# Only keys with a dedicated, dedicated schema node are wired.
OS_SPECIAL_INPUTS: dict[str, tuple[str, str, str | None]] = {
    "lottery_115bb": ("5BB", "LtryPzzlChrgblUs115BB", "IncFrmLottery"),
    "online_games_115bbj": ("5BBJ", "IncChrgblUs115BBJ", "IncFrmOnGames"),
    "unexplained_115bbe": ("5BBE", "IncChrgblUs115BBE", None),
}

SURCHARGE_115BBE_RATE = 0.25  # flat surcharge on tax computed u/s 115BBE

UNSUPPORTED_NOTE = (
    "other special-rate codes (115A-family, FII/FPI 115AD, 115BBF/BBG, "
    "115BBH (VDA — needs Schedule VDA), 115E, DTAA/PTI rows) are validated "
    "but not computed in this version; use a computation prepared with "
    "those schedules and import it, or contact support")


def os_special_rows(data: dict) -> list[dict]:
    """Read ``data['other_sources_special']`` into structured rows:
    ``[{key, code, label, amount, rate, tax}]``; unknown keys raise DataError."""
    src = data.get("other_sources_special") or {}
    rows = []
    for key, amount in src.items():
        if key not in OS_SPECIAL_INPUTS:
            raise DataError(
                f"other_sources_special.{key}: unknown special-rate income key "
                f"'{key}'. Supported: {', '.join(sorted(OS_SPECIAL_INPUTS))}. "
                + ("For " + UNSUPPORTED_NOTE if key.startswith("5") or "_" in key
                   else UNSUPPORTED_NOTE.capitalize() + "."))
        amt = parse_float(amount, field=f"Special-rate income ({key})",
                          path=f"other_sources_special.{key}")
        if amt <= 0:
            continue
        code = OS_SPECIAL_INPUTS[key][0]
        label, rate = SEC_CODES[code]
        rows.append({"key": key, "code": code, "label": label,
                     "amount": int(round(amt)), "rate": rate,
                     "tax": int(round(amt * rate / 100))})
    return rows


def os_special_totals(rows: list[dict]) -> tuple[int, int]:
    return (sum(r["amount"] for r in rows), sum(r["tax"] for r in rows))
