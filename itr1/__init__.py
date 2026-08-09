"""ITR-1 (Sahaj) return preparation package for Assessment Year 2026-27.

Built directly from the official CBDT artifacts in this repository:

* ``ITR-1_2026_Main_V1.1.json``                          - the official JSON schema
* ``CBDT_e-Filing_ITR 1_Validation Rules_AY 2026-27.pdf`` - the validation rules
* ``ITR1_AY_26-27_V1.2.xlsm``                            - the department utility (reference)

The package produces an e-filing JSON that conforms to the official schema,
computes the tax liability under both the new (default, section 115BAC) and
the old tax regimes, and runs the CBDT Category A / B / D validations.
"""

__version__ = "1.0.0"

ASSESSMENT_YEAR = "2026-27"
PREVIOUS_YEAR = "2025-26"

# Software identity written into CreationInfo of the generated JSON.
# NOTE: A production software provider must substitute the SW registration
# number issued by the e-Filing project (pattern: SW + 8 digits).
SW_CREATED_BY = "SW00000001"
SW_VERSION_NO = "1.0"

# ---------------------------------------------------------------------------
# Constants extracted from the official CBDT Excel preparation utilities
# that are checked into this repository (the VBA sources are dumped under
# ``extracted/xlsm/<FORM>/vba/`` — see GenerateJson*.bas, Form01Header() and
# the getSW*/getForm* getters).
#
# ``SW_CREATED_BY`` above stays this software's own identity; the values
# below track the *department utility* so the rest of the identity block
# (SWVersionNo, form descriptions, city, schema/form versions, Digest="-")
# mirrors exactly what the official utility emits for AY 2026-27.
# ---------------------------------------------------------------------------
UTILITY_SW_CREATED_BY = "SW90002627"  # all four utilities, AY 2026-27
UTILITY_INTERMEDIARY_CITY = "Delhi"
# Latest uncommented SWVersionNo in each utility's VBA (30/07/2026 build):
UTILITY_SW_VERSION_BY_FORM = {"ITR1": "R3", "ITR2": "R4", "ITR3": "R4", "ITR4": "R2"}
# Form_ITRn descriptions exactly as the official utilities write them.
# NOTE: the ITR-4 utility's own wording is kept verbatim (including its
# original spelling); the ITR-3 utility ships with a copy-paste bug (it
# reuses the ITR-1 wording) which we do NOT replicate.
UTILITY_FORM_DESCRIPTION = {
    "ITR1": "For Indls having Income from Salary, Pension, family pension "
            "and Interest",
    "ITR2": "For Individuals and HUFs not having income from profits and "
            "gains of business or profession",
    "ITR3": "For Individuals and HUFs having income from profits and gains "
            "of business or profession",
    "ITR4": "For Individuals having Income From Presemptive Business",
}
