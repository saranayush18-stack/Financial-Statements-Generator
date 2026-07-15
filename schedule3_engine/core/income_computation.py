"""
Computation of Total Income and Tax -- Income-tax Act, 2025 (Tax Year 2026-27).

Design principles:
1.  AUTO FIRST: every line is computed automatically from the Schedule III
    statements (PBT, book depreciation, provisions) and the IT depreciation
    schedule.
2.  MANUAL OVERRIDE EVERYWHERE: every AdjustmentLine carries an optional
    `override` -- when set, it wins over the auto amount. The UI exposes this
    per line, matching how a practitioner actually finalizes a computation.
3.  TRANSPARENT: the object keeps both auto and final values so the export
    can show what was overridden.

Structure mirrors the Book1 'Tax Comp' sheet:
    Net Profit as per P&L
      Add: inadmissible expenses (book depreciation, s.37 disallowances,
           penalties, 43B-type unpaid liabilities, TDS-default disallowance)
      Less: admissible (IT depreciation)
    = Business income
      Less: b/f business loss set-off (capped)
      Less: unabsorbed depreciation set-off (capped)
    = Taxable income
      Tax at regime rate/slabs -> surcharge (with marginal relief) -> cess
      Less: TDS / advance tax / self-assessment tax
      Add: interest for default (old 234A/B/C pattern)
    = Net payable / (refund)
"""
from __future__ import annotations

from dataclasses import dataclass, field

from data.tax_config import (
    EntityType, TaxRegime,
    COMPANY_NORMAL_RATE_HIGH, COMPANY_NORMAL_RATE_LOW, COMPANY_TURNOVER_THRESHOLD_FOR_25,
    COMPANY_CONCESSIONAL_RATE, COMPANY_NEW_MFG_RATE, FIRM_RATE,
    COMPANY_SURCHARGE_SLABS, COMPANY_CONCESSIONAL_SURCHARGE, FIRM_SURCHARGE_SLABS,
    HEALTH_EDUCATION_CESS, IND_NEW_REGIME_SLABS, IND_OLD_REGIME_SLABS,
    IND_NEW_REGIME_REBATE_LIMIT, IND_NEW_REGIME_REBATE_MAX,
    IND_OLD_REGIME_REBATE_LIMIT, IND_OLD_REGIME_REBATE_MAX,
    IND_SURCHARGE_SLABS, IND_NEW_REGIME_SURCHARGE_CAP,
    INTEREST_RATE_PER_MONTH,
)
from core.statement_generator import ProfitAndLoss
from core.it_depreciation import ITDepreciationSchedule


@dataclass
class AdjustmentLine:
    """One computation line. `auto` is engine-computed; `override`, when not
    None, replaces it. `final` is what enters the computation."""
    code: str
    label: str
    auto: float = 0.0
    override: float | None = None
    section_ref: str = ""
    note: str = ""

    @property
    def final(self) -> float:
        return round(self.override if self.override is not None else self.auto, 2)

    @property
    def is_overridden(self) -> bool:
        return self.override is not None


@dataclass
class TaxComputation:
    entity_type: EntityType
    regime: TaxRegime
    net_profit_as_per_pnl: float

    additions: list[AdjustmentLine] = field(default_factory=list)
    deductions: list[AdjustmentLine] = field(default_factory=list)

    brought_forward_business_loss: float = 0.0
    unabsorbed_depreciation_bf: float = 0.0

    # Company-normal-regime turnover test for 25% vs 30%
    turnover_reference_year: float = 0.0

    # Prepaid taxes (user inputs, overridable by nature)
    tds_credit: float = 0.0
    tcs_credit: float = 0.0
    advance_tax_paid: float = 0.0
    self_assessment_tax_paid: float = 0.0

    # Interest (auto-simplified; each overridable)
    interest_234a: AdjustmentLine = field(default_factory=lambda: AdjustmentLine(
        "INT_A", "Interest for late filing of return (old 234A pattern)", 0.0,
        section_ref="IT Act 2025 - interest on late return"))
    interest_234b: AdjustmentLine = field(default_factory=lambda: AdjustmentLine(
        "INT_B", "Interest for default in advance tax (old 234B pattern)", 0.0,
        section_ref="IT Act 2025 - advance tax default"))
    interest_234c: AdjustmentLine = field(default_factory=lambda: AdjustmentLine(
        "INT_C", "Interest for deferment of advance tax (old 234C pattern)", 0.0,
        section_ref="IT Act 2025 - advance tax deferment"))

    # ------------------------------------------------------------------
    @property
    def total_additions(self) -> float:
        return round(sum(l.final for l in self.additions), 2)

    @property
    def total_deductions(self) -> float:
        return round(sum(l.final for l in self.deductions), 2)

    @property
    def business_income(self) -> float:
        return round(self.net_profit_as_per_pnl + self.total_additions - self.total_deductions, 2)

    @property
    def loss_set_off(self) -> float:
        """b/f business loss set-off, capped at available business income
        (mirrors Book1's =MIN(F26, CFL!F11))."""
        return round(min(max(self.business_income, 0.0), max(self.brought_forward_business_loss, 0.0)), 2)

    @property
    def unabsorbed_dep_set_off(self) -> float:
        remaining = max(self.business_income - self.loss_set_off, 0.0)
        return round(min(remaining, max(self.unabsorbed_depreciation_bf, 0.0)), 2)

    @property
    def taxable_income(self) -> float:
        return round(self.business_income - self.loss_set_off - self.unabsorbed_dep_set_off, 2)

    # ------------------------------------------------------------------
    def _base_rate_company(self) -> float:
        if self.regime == TaxRegime.COMPANY_CONCESSIONAL_22:
            return COMPANY_CONCESSIONAL_RATE
        if self.regime == TaxRegime.COMPANY_NEW_MFG_15:
            return COMPANY_NEW_MFG_RATE
        if self.turnover_reference_year and self.turnover_reference_year <= COMPANY_TURNOVER_THRESHOLD_FOR_25:
            return COMPANY_NORMAL_RATE_LOW
        return COMPANY_NORMAL_RATE_HIGH

    def _slab_tax(self, slabs: list[tuple[float | None, float]], income: float) -> float:
        tax, lower = 0.0, 0.0
        for upper, rate in slabs:
            if upper is None:
                tax += max(income - lower, 0.0) * rate
                break
            span = min(income, upper) - lower
            if span > 0:
                tax += span * rate
            lower = upper
            if income <= upper:
                break
        return round(tax, 2)

    @property
    def tax_before_surcharge(self) -> float:
        ti = max(self.taxable_income, 0.0)
        if ti == 0:
            return 0.0
        if self.entity_type == EntityType.COMPANY:
            return round(ti * self._base_rate_company(), 2)
        if self.entity_type == EntityType.FIRM_LLP:
            return round(ti * FIRM_RATE, 2)
        # Individual / HUF
        slabs = IND_NEW_REGIME_SLABS if self.regime == TaxRegime.IND_NEW_REGIME else IND_OLD_REGIME_SLABS
        tax = self._slab_tax(slabs, ti)
        # Rebate (old 87A pattern)
        if self.regime == TaxRegime.IND_NEW_REGIME and ti <= IND_NEW_REGIME_REBATE_LIMIT:
            tax = max(tax - IND_NEW_REGIME_REBATE_MAX, 0.0)
        elif self.regime == TaxRegime.IND_OLD_REGIME and ti <= IND_OLD_REGIME_REBATE_LIMIT:
            tax = max(tax - IND_OLD_REGIME_REBATE_MAX, 0.0)
        return round(tax, 2)

    @property
    def surcharge_rate(self) -> float:
        ti = max(self.taxable_income, 0.0)
        if self.entity_type == EntityType.COMPANY:
            if self.regime in (TaxRegime.COMPANY_CONCESSIONAL_22, TaxRegime.COMPANY_NEW_MFG_15):
                return COMPANY_CONCESSIONAL_SURCHARGE
            for threshold, rate in COMPANY_SURCHARGE_SLABS:
                if ti > threshold:
                    return rate
            return 0.0
        if self.entity_type == EntityType.FIRM_LLP:
            for threshold, rate in FIRM_SURCHARGE_SLABS:
                if ti > threshold:
                    return rate
            return 0.0
        for threshold, rate in IND_SURCHARGE_SLABS:
            if ti > threshold:
                if self.regime == TaxRegime.IND_NEW_REGIME:
                    return min(rate, IND_NEW_REGIME_SURCHARGE_CAP)
                return rate
        return 0.0

    @property
    def surcharge(self) -> float:
        raw = round(self.tax_before_surcharge * self.surcharge_rate, 2)
        return self._marginal_relief_applied(raw)

    def _marginal_relief_applied(self, raw_surcharge: float) -> float:
        """Surcharge cannot make total tax exceed tax-at-threshold plus the
        income above the threshold. Simplified single-threshold relief."""
        ti = max(self.taxable_income, 0.0)
        slabs = (COMPANY_SURCHARGE_SLABS if self.entity_type == EntityType.COMPANY
                 and self.regime == TaxRegime.COMPANY_NORMAL
                 else FIRM_SURCHARGE_SLABS if self.entity_type == EntityType.FIRM_LLP
                 else IND_SURCHARGE_SLABS if self.entity_type == EntityType.INDIVIDUAL_HUF
                 else [])
        crossed = None
        for threshold, _rate in sorted(slabs, key=lambda x: x[0]):
            if ti > threshold:
                crossed = threshold
        if crossed is None:
            return raw_surcharge
        saved = TaxComputation(
            entity_type=self.entity_type, regime=self.regime,
            net_profit_as_per_pnl=crossed,  # trick: taxable_income == crossed
            turnover_reference_year=self.turnover_reference_year,
        )
        tax_at_threshold = saved.tax_before_surcharge
        cap = tax_at_threshold + (ti - crossed)
        total = self.tax_before_surcharge + raw_surcharge
        if total > cap:
            return round(max(cap - self.tax_before_surcharge, 0.0), 2)
        return raw_surcharge

    @property
    def cess(self) -> float:
        return round((self.tax_before_surcharge + self.surcharge) * HEALTH_EDUCATION_CESS, 2)

    @property
    def total_tax_liability(self) -> float:
        return round(self.tax_before_surcharge + self.surcharge + self.cess, 2)

    @property
    def prepaid_taxes(self) -> float:
        return round(self.tds_credit + self.tcs_credit + self.advance_tax_paid
                     + self.self_assessment_tax_paid, 2)

    @property
    def total_interest(self) -> float:
        return round(self.interest_234a.final + self.interest_234b.final
                     + self.interest_234c.final, 2)

    @property
    def net_payable(self) -> float:
        """Rounded to nearest Rs. 10 like Book1's =ROUND(...,-1)."""
        raw = self.total_tax_liability - self.prepaid_taxes + self.total_interest
        return round(raw, -1)

    # ------------------------------------------------------------------
    def set_override(self, code: str, amount: float | None):
        for line in self.additions + self.deductions + [
                self.interest_234a, self.interest_234b, self.interest_234c]:
            if line.code == code:
                line.override = amount
                return
        raise KeyError(f"No adjustment line with code {code}")

    def auto_advance_tax_interest(self, months_since_year_end_to_filing: int = 0):
        """Simplified old-234B-pattern interest: 1% per month on the shortfall
        of advance tax below 90% of assessed tax, from 1 April to filing.
        Overridable; a practitioner will refine per actual installment data."""
        assessed = self.total_tax_liability
        shortfall = max(assessed - self.advance_tax_paid - self.tds_credit - self.tcs_credit, 0.0)
        if assessed > 0 and (self.advance_tax_paid + self.tds_credit) < 0.9 * assessed and shortfall > 0:
            months = max(months_since_year_end_to_filing, 1)
            self.interest_234b.auto = round(shortfall * INTEREST_RATE_PER_MONTH * months, 0)
        else:
            self.interest_234b.auto = 0.0


def build_standard_computation(
    entity_type: EntityType,
    regime: TaxRegime,
    pnl: ProfitAndLoss,
    it_dep: ITDepreciationSchedule,
    gratuity_provision_unpaid: float = 0.0,
    leave_encashment_unpaid: float = 0.0,
) -> TaxComputation:
    """Seed the computation with the standard adjustment lines, all
    auto-computed from the financial statements and all overridable."""
    book_depreciation = 0.0
    for e in pnl.expenses:
        if e.sub_head == "Depreciation and Amortization Expense":
            book_depreciation = e.current_year

    comp = TaxComputation(
        entity_type=entity_type, regime=regime,
        net_profit_as_per_pnl=pnl.profit_before_tax_cy,
    )
    comp.additions = [
        AdjustmentLine("ADD_BOOK_DEP", "Depreciation as per books (Companies Act, 2013)",
                       auto=book_depreciation, section_ref="Book-tax timing difference"),
        AdjustmentLine("ADD_S37", "Disallowance u/s 37 (penalties, TDS interest, personal exp.)",
                       auto=0.0, section_ref="s.37, IT Act 2025",
                       note="Enter amounts identified from ledger scrutiny"),
        AdjustmentLine("ADD_TDS_DEFAULT", "Disallowance for TDS default (old 40(a)(ia) pattern)",
                       auto=0.0, section_ref="s.35(b), IT Act 2025"),
        AdjustmentLine("ADD_43B", "Unpaid statutory liabilities (old 43B pattern)",
                       auto=0.0, section_ref="payment-basis deductions, IT Act 2025"),
        AdjustmentLine("ADD_GRATUITY", "Provision for gratuity (unpaid, unapproved fund)",
                       auto=gratuity_provision_unpaid, section_ref="deductible on payment"),
        AdjustmentLine("ADD_LEAVE_ENC", "Provision for leave encashment (unpaid)",
                       auto=leave_encashment_unpaid, section_ref="deductible on payment"),
    ]
    comp.deductions = [
        AdjustmentLine("LESS_IT_DEP", "Depreciation as per Income-tax Act (block WDV)",
                       auto=it_dep.total_depreciation, section_ref="s.33, IT Act 2025 (verify)"),
    ]
    return comp
