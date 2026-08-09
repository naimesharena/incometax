"""Tests for field-identifying input errors (DataError) and tolerant parsing."""
import copy
import json

import pytest

from itr1.builder import build_return
from itr1.errors import DataError, parse_float, parse_int


# A state shaped exactly like the web wizard's blank form (the payload that
# used to crash every API call with "invalid literal for int() with base 10: ''").
def blank_browser_state():
    return {
        "personal": {"first_name": "", "middle_name": "", "last_name": "", "pan": "",
                     "dob": "", "aadhaar": "", "employer_category": "OTH",
                     "address": {"residence_no": "", "residence_name": "", "road_or_street": "",
                                 "locality_or_area": "", "city_or_town_or_district": "",
                                 "state_code": "", "pin_code": "", "country_code": "91",
                                 "mobile_no": "", "email_address": "", "country_code_mobile": "91",
                                 "mobile_no_sec": "", "email_address_sec": ""},
                     "secondary_address_same": True, "alternate_address": {}},
        "filing": {"return_section": 11, "opt_out_new_regime": False,
                   "filing_date": "2026-08-02"},
        "salary": {"salary_17_1": "", "perquisites_17_2": "", "profits_17_3": "",
                   "exempt_allowances": [], "hra": None, "entertainment_allowance": "",
                   "professional_tax": "", "standard_deduction_override": ""},
        "house_properties": [], "other_sources": [],
        "ltcg112a": {"sale_consideration": "", "cost_of_acquisition": ""},
        "exempt_income": [], "deductions": {},
        "tax_paid": {"tds_on_salary": [], "tds_others": [], "tds3": [], "tcs": [],
                     "challans": []},
        "relief_89": "",
        "bank_accounts": [{"ifsc": "", "bank_name": "", "account_no": "",
                           "account_type": "SB", "use_for_refund": True}],
        "verification": {"name": "", "father_name": "", "pan": "", "place": "",
                         "capacity": "S"},
        "interest_override": {},
    }


class TestBlankBrowserState:
    def test_blank_state_builds_without_crash(self):
        res = build_return(blank_browser_state())
        assert res["report"]["total_income"] == 0

    def test_blank_state_summary_endpoint_200(self):
        from itr1.webapp import create_app
        client = create_app().test_client()
        resp = client.post("/api/summary", json=blank_browser_state())
        assert resp.status_code == 200, resp.get_json()

    def test_std_override_empty_string_ok(self):
        data = blank_browser_state()
        data["salary"]["salary_17_1"] = 900000
        rep = build_return(data)["report"]
        assert rep["gross_total_income_incl_ltcg"] == 900000 - 75000

    def test_interest_override_empty_string_ignored(self):
        data = blank_browser_state()
        data["interest_override"] = {"IntrstPayUs234A": "", "LateFilingFee234F": ""}
        build_return(data)  # must not crash


class TestFieldIdentifyingErrors:
    def _expect(self, data, field_part, path_part):
        with pytest.raises(DataError) as ei:
            build_return(data)
        exc = ei.value
        assert field_part in exc.field, f"{exc.field!r} should mention {field_part!r}"
        assert path_part in exc.path, f"{exc.path!r} should mention {path_part!r}"
        assert exc.field in str(exc), "message should lead with the field name"
        assert "invalid literal for int()" not in str(exc)

    def test_std_override_garbage(self):
        d = blank_browser_state()
        d["salary"]["standard_deduction_override"] = "abc"
        self._expect(d, "Standard deduction override", "salary.standard_deduction_override")

    def test_salary_garbage(self):
        d = blank_browser_state()
        d["salary"]["salary_17_1"] = "twelve lakh"
        self._expect(d, "Salary as per section 17(1)", "salary.salary_17_1")

    def test_exempt_allowance_row_garbage(self):
        d = blank_browser_state()
        d["salary"]["exempt_allowances"] = [{"section": "10(5)", "amount": "5k"}]
        self._expect(d, "Exempt allowance row 1", "salary.exempt_allowances.0.amount")

    def test_coowner_share_garbage(self):
        d = blank_browser_state()
        d["house_properties"] = [{"occupancy": "S", "co_owned": True,
                                  "co_owners": [{"name": "Wife", "share_pct": "half"}]}]
        self._expect(d, "House property 1", "house_properties.0.share_pct")

    def test_challan_amount_garbage(self):
        d = blank_browser_state()
        d["tax_paid"]["challans"] = [{"bsr_code": "0001234", "date": "2025-09-10",
                                      "challan_no": "12", "amount": "ten thousand"}]
        self._expect(d, "challan 1", "tax_paid.challans.0.amount")

    def test_tds_row_garbage(self):
        d = blank_browser_state()
        d["tax_paid"]["tds_on_salary"] = [{"tan": "AAAA00000A", "employer_name": "E",
                                           "tds": "1lakh"}]
        self._expect(d, "TDS on salary row 1", "tax_paid.tds_on_salary.0.tds")

    def test_80c_item_garbage(self):
        d = blank_browser_state()
        d["deductions"] = {"80c": {"items": [{"identification_no": "LIC", "amount": "12%"}]}}
        self._expect(d, "80C investment row 1", "deductions.80c.items.0.amount")

    def test_interest_override_garbage(self):
        d = blank_browser_state()
        d["interest_override"] = {"IntrstPayUs234B": "high"}
        self._expect(d, "234B", "interest_override.IntrstPayUs234B")

    def test_relief89_garbage(self):
        d = blank_browser_state()
        d["relief_89"] = "?"
        self._expect(d, "Relief u/s 89", "relief_89")

    def test_api_error_carries_field_and_path(self):
        from itr1.webapp import create_app
        client = create_app().test_client()
        d = blank_browser_state()
        d["salary"]["standard_deduction_override"] = "abc"
        resp = client.post("/api/summary", json=d)
        assert resp.status_code == 400
        payload = resp.get_json()
        assert payload["path"] == "salary.standard_deduction_override"
        assert "Standard deduction override" in payload["field"]
        assert payload["error"].startswith(payload["field"])

    def test_dataerror_is_valueerror_for_backward_compat(self):
        try:
            build_return({"salary": {"salary_17_1": "zzz"}})
            raise AssertionError("should have raised")
        except ValueError:
            pass


class TestTolerantParsing:
    def test_indian_grouped_commas(self):
        rep = build_return({"salary": {"salary_17_1": "12,00,000"}})["report"]
        assert rep["gross_total_income_incl_ltcg"] == 1200000 - 75000

    def test_western_grouped_commas(self):
        assert parse_int("1,200,000") == 1200000

    def test_decimal_text_rounded_half_up(self):
        assert parse_int("12000.50") == 12001
        assert parse_int("12000.49") == 12000

    def test_whitespace_and_currency_marks(self):
        assert parse_int("  75000 ") == 75000
        assert parse_int("₹ 75,000") == 75000
        assert parse_int("Rs.1000") == 1000

    def test_blank_defaults(self):
        assert parse_int("") == 0
        assert parse_int(None) == 0
        assert parse_int("", default=None) is None
        assert parse_float("") == 0.0
        assert parse_float("50") == 50.0

    def test_nan_rejected_with_field(self):
        with pytest.raises(DataError):
            parse_int(float("nan"), field="Test field")

    def test_negative_allowed(self):
        assert parse_int("-5000") == -5000

    @pytest.mark.parametrize("bad", ["abc", "5k", "12%", "1,0", "3..4", "--5"])
    def test_garbage_rejected(self, bad):
        with pytest.raises(DataError):
            parse_int(bad, field="X")
