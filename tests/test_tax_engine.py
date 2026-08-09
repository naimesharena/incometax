"""Unit tests for the AY 2026-27 tax computation engine."""

import datetime as dt

from itr1 import tax_engine as T

YOUNG = "1991-06-15"
SENIOR = "1960-01-01"
SUPER_SENIOR = "1940-01-01"

BASE = dict(ltcg112a=0, section89=0, filing_date="2026-07-15",
            return_section=11, tds=0, tcs=0, advance_tax_challans=[])


def ct(total_income, opt_out, dob=YOUNG, **kw):
    cfg = dict(BASE, **kw)
    return T.compute_tax(total_income=total_income, opt_out_new_regime=opt_out,
                         dob=dob, **cfg)


# ---------------------------------------------------------------- slabs
class TestSlabs:
    def test_new_regime_at_15lakh(self):
        res = ct(1_500_000, opt_out=False)
        assert res["TotalTaxPayable"] == 105_000       # 60k + 15% of 3L
        assert res["EducationCess"] == 4_200
        assert res["GrossTaxLiability"] == 109_200

    def test_new_regime_at_25lakh(self):
        res = ct(2_500_000, opt_out=False)
        assert res["TotalTaxPayable"] == 3_30_000
        assert res["GrossTaxLiability"] == 3_30_000 * 1.04

    def test_old_regime_general_at_10lakh(self):
        res = ct(1_000_000, opt_out=True)
        assert res["TotalTaxPayable"] == 112_500       # 12.5k + 20% of 5L
        assert res["GrossTaxLiability"] == 117_000

    def test_old_regime_senior_at_10lakh(self):
        res = ct(1_000_000, opt_out=True, dob=SENIOR)
        assert res["TotalTaxPayable"] == 110_000       # 10k + 20% of 5L

    def test_old_regime_super_senior_at_10lakh(self):
        res = ct(1_000_000, opt_out=True, dob=SUPER_SENIOR)
        assert res["TotalTaxPayable"] == 100_000       # 20% of 5L


# ---------------------------------------------------------------- ages
class TestAge:
    def test_senior_boundary(self):
        assert T.is_senior_citizen("1966-04-01")
        assert not T.is_senior_citizen("1966-04-02")

    def test_super_senior_boundary(self):
        assert T.is_super_senior_citizen("1946-04-01")
        assert not T.is_super_senior_citizen("1946-04-02")


# ---------------------------------------------------------------- 87A
class TestRebate:
    def test_new_12lakh_zero_tax(self):
        res = ct(1_200_000, opt_out=False)
        assert res["TotalTaxPayable"] == 60_000
        assert res["Rebate87A"] == 60_000
        assert res["GrossTaxLiability"] == 0

    def test_new_1275_salary_equivalent(self):
        # 12,75,000 salary - 75,000 SD = 12,00,000 taxable -> zero (tested at TI level)
        res = ct(1_200_000, opt_out=False)
        assert res["TotTaxPlusIntrstPay"] == 0

    def test_marginal_relief_at_1201000(self):
        res = ct(1_201_000, opt_out=False)
        # slab tax = 60,150; marginal cap = 1,000
        assert res["Rebate87A"] == 59_150
        assert res["TaxPayableOnRebate"] == 1_000
        assert res["EducationCess"] == 40

    def test_marginal_relief_ceiling_1270590(self):
        res = ct(1_270_590, opt_out=False)
        assert res["Rebate87A"] == 0 or res["TaxPayableOnRebate"] <= 70_590
        res2 = ct(1_271_000, opt_out=False)
        assert res2["Rebate87A"] == 0

    def test_rebate_never_exceeds_60000(self):
        for inc in range(1_200_000, 1_400_000, 7_777):
            res = ct(inc, opt_out=False)
            assert res["Rebate87A"] <= 60_000

    def test_old_regime_5lakh(self):
        res = ct(500_000, opt_out=True)
        assert res["Rebate87A"] == 12_500
        assert res["GrossTaxLiability"] == 0

    def test_old_regime_above_5lakh_no_rebate(self):
        res = ct(500_100, opt_out=True)
        assert res["Rebate87A"] == 0


# ---------------------------------------------------------------- surcharge
class TestSurcharge:
    def test_surcharge_above_50l_with_marginal_relief(self):
        # only reachable with LTCG 112A (ITR-1 caps normal income at 50L)
        res = ct(5_050_000, opt_out=False)
        assert res["TotalTaxPayable"] > 0
        # capped: tax on 50L + income excess of 50L
        res50 = ct(5_000_000, opt_out=False)
        assert res["TaxPayableOnRebate"] <= res50["TotalTaxPayable"] + 50_000 + 1

    def test_no_surcharge_at_50l(self):
        res = ct(5_000_000, opt_out=False)
        assert res["SurchargeComputed"] == 0


# ---------------------------------------------------------------- LTCG 112A
class TestLTCG:
    def test_ltcg_upto_125l_exempt(self):
        res = ct(total_income=900_000, opt_out=False, **{**BASE, "ltcg112a": 100_000})
        # tax only on 8,00,000 at slab: 5% of 4L = 20,000
        assert res["TotalTaxPayable"] == 20_000

    def test_ltcg_tax_function_boundary(self):
        assert T.ltcg_tax_112a(125_000) == 0
        assert T.ltcg_tax_112a(126_000) == 125  # 12.5% of 1,000


# ---------------------------------------------------------------- 234F
class TestFee234F:
    def test_on_time_no_fee(self):
        assert T.fee_234f(600_000, "2026-07-31", 11) == 0

    def test_late_full_fee(self):
        assert T.fee_234f(600_000, "2026-08-02", 12) == 5_000

    def test_late_small_income(self):
        assert T.fee_234f(450_000, "2026-08-02", 12) == 1_000

    def test_fee_234i_revised_after_cutoff(self):
        assert T.fee_234i(600_000, "2027-01-15", 17) == 5_000
        assert T.fee_234i(400_000, "2027-01-15", 17) == 1_000
        assert T.fee_234i(600_000, "2026-12-31", 17) == 0
        assert T.fee_234i(600_000, "2027-01-15", 11) == 0


# ---------------------------------------------------------------- interest
class TestInterest:
    def test_234a_two_months(self):
        # filed 2 months after due date on 1,00,000 unpaid tax
        v = T.interest_234a(100_000, 0, "2026-09-20", 12)
        assert v == 2_000

    def test_234b_below_90_percent(self):
        v = T.interest_234b(100_000, 0, "2026-08-02")
        assert v == T.floor_hundred(100_000) * 0.01 * T._months_between(dt.date(2026, 3, 31), dt.date(2026, 8, 2))

    def test_234b_sufficient_advance_tax(self):
        assert T.interest_234b(100_000, 91_000, "2026-08-02") == 0

    def test_234c_full_deferment(self):
        v = T.interest_234c(100_000, 0, [])
        # 15,000*3 + 45,000*3 + 75,000*3 + 100,000*1 at 1% = 450+1350+2250+1000
        assert v == 5_050

    def test_234c_paid_on_time(self):
        challans = [
            {"date": "2025-06-14", "amount": 15_000},
            {"date": "2025-09-14", "amount": 30_000},
            {"date": "2025-12-14", "amount": 30_000},
            {"date": "2026-03-14", "amount": 25_000},
        ]
        assert T.interest_234c(100_000, 0, challans) == 0

    def test_234c_not_applicable_below_10k(self):
        assert T.interest_234c(9_999, 0, []) == 0


# ---------------------------------------------------------------- rounding
class TestRounding:
    def test_income_round_off_288a(self):
        assert T.round_off_income(1_234_564) == 1_234_560
        assert T.round_off_income(1_234_565) == 1_234_570

    def test_tax_round_off_288b(self):
        assert T.round_off_tax(16_224) == 16_220
