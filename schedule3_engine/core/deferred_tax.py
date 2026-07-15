"""
Deferred Tax (AS 22 pattern for non-Ind AS companies), mirroring Book1's
'Deferred Tax' sheet:

    Particulars | Opening DTA/(DTL) | As per Income Tax | As per Books |
    Timing Difference | Current-year DTA/(DTL) | Closing DTA/(DTL)

Convention: timing_difference = books - income_tax.
    * Fixed assets: book WDV > IT WDV  -> positive difference -> DTA
      (book depreciation will exceed IT depreciation in future years).
    * Provisions disallowed now, deductible on payment -> DTA.
A negative closing figure is a Deferred Tax Liability.

The effective rate follows the chosen regime (e.g. 22% + 10% surcharge + 4%
cess = 25.168%, exactly the 0.25168 hard-coded in Book1).
"""
from __future__ import annotations

from dataclasses import dataclass, field

from data.tax_config import HEALTH_EDUCATION_CESS


@dataclass
class DeferredTaxItem:
    particulars: str
    opening_dta_dtl: float = 0.0
    as_per_income_tax: float = 0.0
    as_per_books: float = 0.0
    override_timing_difference: float | None = None

    @property
    def timing_difference(self) -> float:
        if self.override_timing_difference is not None:
            return round(self.override_timing_difference, 2)
        return round(self.as_per_books - self.as_per_income_tax, 2)


@dataclass
class DeferredTaxComputation:
    effective_rate: float
    items: list[DeferredTaxItem] = field(default_factory=list)

    def closing_for(self, item: DeferredTaxItem) -> float:
        return round(item.timing_difference * self.effective_rate, 2)

    def movement_for(self, item: DeferredTaxItem) -> float:
        return round(self.closing_for(item) - item.opening_dta_dtl, 2)

    @property
    def total_opening(self) -> float:
        return round(sum(i.opening_dta_dtl for i in self.items), 2)

    @property
    def total_closing(self) -> float:
        return round(sum(self.closing_for(i) for i in self.items), 2)

    @property
    def total_movement(self) -> float:
        """Charge/(credit) to P&L for the year. Positive = DTA created
        (credit to P&L, reduces tax expense)."""
        return round(self.total_closing - self.total_opening, 2)


def effective_tax_rate(base_rate: float, surcharge_rate: float) -> float:
    """e.g. 0.22 base, 0.10 surcharge -> 0.22*1.10*1.04 = 0.25168"""
    return round(base_rate * (1 + surcharge_rate) * (1 + HEALTH_EDUCATION_CESS), 6)


def build_standard_deferred_tax(
    effective_rate: float,
    it_closing_wdv: float,
    book_closing_wdv: float,
    gratuity_provision: float = 0.0,
    leave_encashment_provision: float = 0.0,
    tds_default_disallowance: float = 0.0,
    opening_balances: dict[str, float] | None = None,
) -> DeferredTaxComputation:
    ob = opening_balances or {}
    comp = DeferredTaxComputation(effective_rate=effective_rate)
    comp.items = [
        DeferredTaxItem("Fixed Assets (WDV: books vs Income-tax)",
                        opening_dta_dtl=ob.get("Fixed Assets", 0.0),
                        as_per_income_tax=it_closing_wdv, as_per_books=book_closing_wdv),
        DeferredTaxItem("Provision for Gratuity (deductible on payment)",
                        opening_dta_dtl=ob.get("Gratuity", 0.0),
                        as_per_income_tax=0.0, as_per_books=gratuity_provision),
        DeferredTaxItem("Provision for Leave Encashment (deductible on payment)",
                        opening_dta_dtl=ob.get("Leave Encashment", 0.0),
                        as_per_income_tax=0.0, as_per_books=leave_encashment_provision),
        DeferredTaxItem("Disallowance for TDS default (deductible on remittance)",
                        opening_dta_dtl=ob.get("TDS Default", 0.0),
                        as_per_income_tax=0.0, as_per_books=tds_default_disallowance),
    ]
    return comp
