"""Statutory constants for AY 2026-27 (Previous Year 2025-26) - ITR-1 (Sahaj).

Sources:
  * Finance Act 2025 (new-regime slabs for FY 2025-26, section 115BAC)
  * ITR-1 JSON schema V1.1 (caps embedded on individual fields)
  * CBDT ITR-1 Validation Rules for AY 2026-27 (rule numbers quoted inline)
"""

# ---------------------------------------------------------------------------
# Tax regime slabs: list of (lower_exclusive_limit, rate_percent, cumulative_tax)
# ---------------------------------------------------------------------------

# New regime (default, u/s 115BAC) for FY 2025-26 / AY 2026-27
NEW_REGIME_SLABS = [
    (0, 0.0),
    (400_000, 0.05),
    (800_000, 0.10),
    (1_200_000, 0.15),
    (1_600_000, 0.20),
    (2_000_000, 0.25),
    (2_400_000, 0.30),
]

# Old regime - individual resident below 60 years
OLD_REGIME_SLABS_GENERAL = [
    (0, 0.0),
    (250_000, 0.05),
    (500_000, 0.20),
    (1_000_000, 0.30),
]

# Old regime - senior citizen (60 years or more but below 80 years)
OLD_REGIME_SLABS_SENIOR = [
    (0, 0.0),
    (300_000, 0.05),
    (500_000, 0.20),
    (1_000_000, 0.30),
]

# Old regime - super senior citizen (80 years or more)
OLD_REGIME_SLABS_SUPER_SENIOR = [
    (0, 0.0),
    (500_000, 0.20),
    (1_000_000, 0.30),
]

# Age boundaries as per the CBDT validation rules (rules 13 and 15):
# a senior citizen for AY 2026-27 is born on or before 01-04-1966.
SENIOR_CITIZEN_DOB_ON_OR_BEFORE = "1966-04-01"
SUPER_SENIOR_DOB_ON_OR_BEFORE = "1946-04-01"

# ---------------------------------------------------------------------------
# Rebate u/s 87A
# ---------------------------------------------------------------------------
REBATE_87A_NEW_MAX = 60_000            # schema caps Rebate87A at 60,000
REBATE_87A_NEW_INCOME_LIMIT = 1_200_000
# Marginal-relief boundary per CBDT rule 191 (Rs. 12,70,590)
REBATE_87A_NEW_MR_CEILING = 1_270_590

REBATE_87A_OLD_MAX = 12_500            # CBDT rule 192
REBATE_87A_OLD_INCOME_LIMIT = 500_000  # CBDT rule 23 (total income incl. LTCG)

# ---------------------------------------------------------------------------
# Surcharge thresholds (only the 10% slab can ever trigger in ITR-1, total
# income is capped at Rs. 50 lakh + Rs. 1.25 lakh LTCG 112A)
# ---------------------------------------------------------------------------
SURCHARGE = [
    # (income_above, rate, cap_on_max_applicable_rate_new_regime, cap_old_regime)
    (5_000_000, 0.10),
    (10_000_000, 0.15),
    (20_000_000, 0.25),
    (50_000_000, 0.37),  # 37% applies only to the old regime; new caps at 25%
]
SURCHARGE_MAX_RATE_NEW_REGIME = 0.25

HEALTH_EDUCATION_CESS = 0.04

# ---------------------------------------------------------------------------
# Salary deductions u/s 16
# ---------------------------------------------------------------------------
STD_DEDUCTION_16IA_NEW = 75_000  # CBDT rule 215; schema max on DeductionUs16ia
STD_DEDUCTION_16IA_OLD = 50_000  # CBDT rule 112
ENTERTAINMENT_ALLOW_16II_MAX = 5_000   # rules 57/58; schema max 5,000
PROFESSIONAL_TAX_16III_MAX = 5_000     # schema max 5,000

# Family pension deduction u/s 57(iia): lower of 1/3rd of family pension or cap
FAMILY_PENSION_DED_CAP_OLD = 15_000    # rule 54
FAMILY_PENSION_DED_CAP_NEW = 25_000    # rule 214

# ---------------------------------------------------------------------------
# House property
# ---------------------------------------------------------------------------
HP_SELF_OCCUPIED_INTEREST_CAP_OLD = 200_000  # rule 48
HP_MAX_PROPERTIES = 2                         # schema: PropertyDetails maxItems=2
HP_STD_DEDUCTION_RATE = 0.30                  # rule 43
HP_ARREARS_TAXABLE_FRACTION = 0.70            # arrears less 30%

# ---------------------------------------------------------------------------
# Chapter VI-A deduction caps
# ---------------------------------------------------------------------------
CAP_80C_CCC_CCD1 = 150_000         # rule 1 (combined)
CAP_80CCD1B = 50_000               # rule 115
CAP_80CCD2_PCT_CG_SG_OLD = 0.14    # rule 120
CAP_80CCD2_PCT_OTHERS_OLD = 0.10   # rule 4
CAP_80CCD2_PCT_NEW = 0.14          # rule 216
CAP_80CCD1_PCT_SALARIED = 0.10     # rule 3
CAP_80CCD1_PCT_GTI_OTHERS = 0.20   # rule 2
CAP_80D_SELF_FAMILY = 25_000       # rule 127
CAP_80D_SELF_FAMILY_SR = 50_000    # rule 130
CAP_80D_PARENTS = 25_000           # rule 132
CAP_80D_PARENTS_SR = 50_000        # rule 134
CAP_80D_PREV_CHECKUP = 5_000       # rule 129 (all fields combined)
CAP_80D_TOTAL = 100_000            # rule 136; schema max
CAP_80DD = 75_000                  # rule 203 (disability)
CAP_80DD_SEVERE = 125_000          # rule 204
CAP_80DDB_NORMAL = 40_000          # rule 7
CAP_80DDB_SR = 100_000             # rule 5
CAP_80EE = 50_000                  # rule 121
CAP_80EEA = 150_000                # rule 122
CAP_80EEB = 150_000                # rule 124
CAP_80GG = 60_000                  # rule 114 (also 25% of adjusted TI)
CAP_80GG_PCT_OF_ADJ_TI = 0.25
CAP_80TTA = 10_000                 # rule 11
CAP_80TTB = 50_000                 # rule 14
CAP_80U = 75_000                   # rule 201
CAP_80U_SEVERE = 125_000           # rule 200
CAP_80CCH_PCT_OF_SALARY = 0.462    # rule 186
CAP_80CCH_SCHEMA_MAX = 288_000     # schema max on AnyOthSec80CCH
CAP_80EE_LOAN_AMOUNT = 3_500_000   # rule 227
CAP_80EEA_STAMP_DUTY_VALUE = 4_500_000  # rule 229; schema max on PropStmpDtyVal

# 80G / 80GGA cash donation limit (rules 88, 143)
CASH_DONATION_LIMIT = 2_000

# ---------------------------------------------------------------------------
# ITR-1 eligibility / field caps taken from the schema
# ---------------------------------------------------------------------------
MAX_AGRICULTURE_INCOME = 5_000          # rule 29
MAX_LTCG_112A = 125_000                 # schema max on LongCap112A; rule 217
LTCG_112A_RATE = 0.125                  # rate above the exempt 1.25 lakh
STCG_111A_RATE = 0.20                   # STCG on equity/equity-MF (STT) post 23-07-2024
LTCG_GENERAL_125_RATE = 0.125           # other LTCG w/o indexation post 23-07-2024
LTCG_INDEXED_20_RATE = 0.20             # LTCG at 20% with indexation (pre-23-07-2024 option)
MAX_TOTAL_INCOME_ITR1 = 5_125_000       # schema max on TotalIncome
MAX_INCOME_EXCL_LTCG = 5_000_000        # rule 117

MAX_VALUE = 99_999_999_999_999

# ---------------------------------------------------------------------------
# Interest / fee sections
# ---------------------------------------------------------------------------
DUE_DATE_FILING = "2026-07-31"          # schema: ItrFilingDueDate pattern
LATE_FEE_234F_SMALL = 1_000             # total income <= 5 lakh  (rule 324 logic)
LATE_FEE_234F_FULL = 5_000              # total income  > 5 lakh
FEE_234I_SMALL = 1_000                  # rule 324: revised return after 31-12-2026
FEE_234I_FULL = 5_000                   # rule 328
REVISED_RETURN_FEE_CUTOFF = "2026-12-31"
INTEREST_234_RATE = 0.01                # 1% per month or part of a month

# Advance tax instalment deadlines for 234C
ADVANCE_TAX_INSTALMENTS = [
    ("2025-06-15", 0.15),
    ("2025-09-15", 0.45),
    ("2025-12-15", 0.75),
    ("2026-03-15", 1.00),
]

# ---------------------------------------------------------------------------
# Filing sections (FilingStatus.ReturnFileSec enum)
# ---------------------------------------------------------------------------
RETURN_SECTIONS = {
    11: "139(1) - On or before due date",
    12: "139(4) - After due date",
    13: "142(1)",
    14: "148",
    16: "153C",
    17: "139(5) - Revised",
    18: "139(9)",
    20: "119(2)(b) - After condonation of delay",
}

EMPLOYER_CATEGORIES = {
    "CGOV": "Central Government",
    "SGOV": "State Government",
    "PSU": "Public Sector Undertaking",
    "PE": "Pensioners - Central Government",
    "PESG": "Pensioners - State Government",
    "PEPS": "Pensioners - Public Sector Undertaking",
    "PEO": "Pensioners - Others",
    "OTH": "Others",
    "NA": "Not Applicable",
}
PENSIONER_CATEGORIES = ("PE", "PESG", "PEPS", "PEO")

STATE_CODES = {
    "01": "Andaman and Nicobar Islands", "02": "Andhra Pradesh",
    "03": "Arunachal Pradesh", "04": "Assam", "05": "Bihar",
    "06": "Chandigarh", "07": "Dadra and Nagar Haveli and Daman and Diu",
    "08": "Delhi", "09": "Goa", "10": "Gujarat", "11": "Haryana",
    "12": "Himachal Pradesh", "13": "Jammu and Kashmir", "14": "Karnataka",
    "15": "Kerala", "16": "Ladakh", "17": "Lakshadweep", "18": "Madhya Pradesh",
    "19": "Maharashtra", "20": "Manipur", "21": "Meghalaya", "22": "Mizoram",
    "23": "Nagaland", "24": "Odisha", "25": "Puducherry", "26": "Punjab",
    "27": "Rajasthan", "28": "Sikkim", "29": "Tamil Nadu", "30": "Telangana",
    "31": "Tripura", "32": "Uttar Pradesh", "33": "Uttarakhand",
    "34": "West Bengal", "35": "Chhattisgarh", "36": "Jharkhand",
    "37": "Andhra Pradesh (New)", "99": "Outside India",
}

# Exempt salary allowances u/s 10 (SalNatureDesc enum)
EXEMPT_ALLOWANCE_SECTIONS = {
    "10(5)": "Leave Travel concession/assistance",
    "10(6)": "Remuneration of official of embassy etc.",
    "10(7)": "Allowances/perquisites paid outside India by Government",
    "10(10)": "Death-cum-retirement gratuity received",
    "10(10A)": "Commuted value of pension received",
    "10(10AA)": "Earned leave encashment on retirement",
    "10(10B)(i)": "Retrenchment compensation - first proviso",
    "10(10B)(ii)": "Retrenchment compensation - approved scheme",
    "10(10C)": "Voluntary retirement/termination compensation",
    "10(10CC)": "Tax paid by employer on non-monetary perquisite",
    "10(13A)": "House Rent Allowance",
    "10(14)(i)": "Prescribed allowances for performance of duties",
    "10(14)(ii)": "Allowances for personal expenses/cost of living",
    "10(14)(i)(115BAC)": "Rule 2BB allowances (new regime only)",
    "10(14)(ii)(115BAC)": "Transport allowance to divyang employee (new regime)",
    "EIC": "Exempt income received by a Judge (SC/HC)",
    "10(17)": "Allowance to MP/MLA/MLC",
}

# Other income nature codes (OthSrcNatureDesc enum)
OTHER_SOURCE_NATURES = {
    "SAV": "Interest from Savings Account",
    "IFD": "Interest from Deposit (Bank/Post Office/Cooperative Society)",
    "TAX": "Interest from Income Tax Refund",
    "FAP": "Family Pension",
    "DIV": "Dividend",
    "10(11)(iP)": "Taxable PF interest - first proviso to sec 10(11)",
    "10(11)(iiP)": "Taxable PF interest - second proviso to sec 10(11)",
    "10(12)(iP)": "Taxable RPF interest - first proviso to sec 10(12)",
    "10(12)(iiP)": "Taxable RPF interest - second proviso to sec 10(12)",
    "OTH": "Any Other",
}
