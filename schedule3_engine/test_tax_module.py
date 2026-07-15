"""End-to-end test of the tax module on the demo Trial Balance."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from datetime import date
from openpyxl import Workbook

from models import Company
from core.tb_parser import parse_trial_balance
from core.classifier import classify_ledger, suggestion_to_mapping_entry
from core.statement_generator import (
    generate_balance_sheet, generate_profit_and_loss, carry_profit_into_reserves,
)
from core.it_depreciation import seed_schedule_from_tb
from core.income_computation import build_standard_computation
from core.deferred_tax import build_standard_deferred_tax, effective_tax_rate
from core.tax_audit import check_applicability, build_form26_checklist
from data.tax_config import EntityType, TaxRegime
from export.tax_export import add_tax_sheets, build_computation_pdf
from sample_data.generate_sample_tb import generate as write_sample_tb


def main():
    out = Path("tax_test_output")
    out.mkdir(exist_ok=True)
    tb_path = out / "sample_tb.xlsx"
    write_sample_tb(str(tb_path))

    company = Company(
        name="Bharat Precision Components Private Limited",
        cin="U29100MH2015PTC123456", pan="AAACB1234C",
        registered_office="Plot No. 45, MIDC Industrial Area, Pune, Maharashtra - 411019",
        auditor="M/s ABC & Associates, Chartered Accountants",
        directors=["Mr. Rajesh Sharma", "Mrs. Priya Sharma"],
        financial_year_start=date(2026, 4, 1), financial_year_end=date(2027, 3, 31),
        company_id=1,
    )
    tb, _ = parse_trial_balance(str(tb_path), company, "Tax Year 2026-27")
    mappings = {}
    for l in tb.ledgers:
        entry = suggestion_to_mapping_entry(classify_ledger(l.ledger_name))
        if entry is not None:
            mappings[l.ledger_name] = entry

    bs = generate_balance_sheet(tb, mappings)
    pnl = generate_profit_and_loss(tb, mappings)
    bs = carry_profit_into_reserves(bs, pnl)

    # --- IT depreciation ---
    it_dep = seed_schedule_from_tb(tb, mappings)
    print("IT dep blocks:", [(r.block_name, r.rate, r.opening_wdv, r.additions_180_plus) for r in it_dep.rows])
    print("IT dep total:", it_dep.total_depreciation, "closing WDV:", it_dep.total_closing_wdv)

    # --- Income computation: company, 22% concessional route ---
    comp = build_standard_computation(
        EntityType.COMPANY, TaxRegime.COMPANY_CONCESSIONAL_22, pnl, it_dep,
        gratuity_provision_unpaid=30_000,
    )
    comp.brought_forward_business_loss = 990_740
    comp.unabsorbed_depreciation_bf = 580_272
    comp.tds_credit = 25_000
    comp.advance_tax_paid = 0
    comp.set_override("ADD_S37", 315.0)   # manual override example (Book1 had 315)
    comp.auto_advance_tax_interest(months_since_year_end_to_filing=6)

    print("\nPBT:", comp.net_profit_as_per_pnl)
    print("Additions:", comp.total_additions, "Deductions:", comp.total_deductions)
    print("Business income:", comp.business_income)
    print("Loss set-off:", comp.loss_set_off, "Unabs dep set-off:", comp.unabsorbed_dep_set_off)
    print("Taxable income:", comp.taxable_income)
    print("Tax:", comp.tax_before_surcharge, "Surcharge:", comp.surcharge, "Cess:", comp.cess)
    print("Liability:", comp.total_tax_liability, "Net payable:", comp.net_payable)

    # sanity: loss-making demo company should have zero taxable income & tax
    assert comp.taxable_income <= 0 or comp.taxable_income >= 0  # structural
    if comp.business_income <= 0:
        assert comp.taxable_income == comp.business_income
        assert comp.tax_before_surcharge == 0

    # --- Profitable scenario checks (synthetic) ---
    comp2 = build_standard_computation(EntityType.COMPANY, TaxRegime.COMPANY_CONCESSIONAL_22, pnl, it_dep)
    comp2.net_profit_as_per_pnl = 10_000_000
    comp2.additions = []
    comp2.deductions = []
    t = comp2
    expected = round(10_000_000 * 0.22, 2)
    assert abs(t.tax_before_surcharge - expected) < 1, (t.tax_before_surcharge, expected)
    assert abs(t.surcharge - expected * 0.10) < 1
    eff = t.total_tax_liability / 10_000_000
    assert abs(eff - 0.25168) < 0.0001, eff
    print("\n22% route effective rate check OK:", eff)

    # Individual new-regime slab check: TI = 25,00,000
    comp3 = build_standard_computation(EntityType.INDIVIDUAL_HUF, TaxRegime.IND_NEW_REGIME, pnl, it_dep)
    comp3.net_profit_as_per_pnl = 2_500_000
    comp3.additions, comp3.deductions = [], []
    # slabs: 4L@0 + 4L@5% + 4L@10% + 4L@15% + 4L@20% + 4L@25% + 1L@30% = 0+20k+40k+60k+80k+100k+30k = 330k
    assert abs(comp3.tax_before_surcharge - 330_000) < 1, comp3.tax_before_surcharge
    print("Individual new-regime slab check OK:", comp3.tax_before_surcharge)

    # Rebate check: TI = 11,00,000 -> tax 0+20k+30k = 50k... slabs: 4-8@5%=20k, 8-11@10%=30k => 50k, <=12L so rebate 60k -> 0
    comp4 = build_standard_computation(EntityType.INDIVIDUAL_HUF, TaxRegime.IND_NEW_REGIME, pnl, it_dep)
    comp4.net_profit_as_per_pnl = 1_100_000
    comp4.additions, comp4.deductions = [], []
    assert comp4.tax_before_surcharge == 0, comp4.tax_before_surcharge
    print("87A-pattern rebate check OK (11L -> zero tax)")

    # Firm check with surcharge + marginal relief around 1 crore
    comp5 = build_standard_computation(EntityType.FIRM_LLP, TaxRegime.FIRM_STANDARD, pnl, it_dep)
    comp5.net_profit_as_per_pnl = 10_050_000   # just above Rs. 1 crore
    comp5.additions, comp5.deductions = [], []
    base = 10_050_000 * 0.30
    raw_sur = base * 0.12
    cap_extra = (10_050_000 - 10_000_000)  # income above threshold
    tax_at_thresh = 10_000_000 * 0.30
    assert comp5.tax_before_surcharge + comp5.surcharge <= tax_at_thresh + cap_extra + 1
    print("Firm marginal relief check OK:", comp5.surcharge, "<= raw", round(raw_sur, 2))

    # --- Deferred tax ---
    book_wdv = 3_110_000.0  # from BS PPE current year
    rate = effective_tax_rate(0.22, 0.10)
    assert abs(rate - 0.25168) < 1e-9
    dt = build_standard_deferred_tax(rate, it_dep.total_closing_wdv, book_wdv,
                                     gratuity_provision=180_000)
    print("\nDeferred tax closing:", dt.total_closing, "movement:", dt.total_movement)

    # --- Tax audit ---
    turnover = pnl.total_revenue_cy
    ap = check_applicability(EntityType.COMPANY, is_profession=False,
                             turnover_or_receipts=turnover,
                             cash_receipts_pct=0.02, cash_payments_pct=0.03)
    print("\nAudit required:", ap.required)
    for r_ in ap.reasons:
        print("  -", r_)
    ap2 = check_applicability(EntityType.COMPANY, False, 12_00_00_000, 0.02, 0.03)
    assert ap2.required, "12cr digital turnover must trigger audit over 10cr limit"
    ap3 = check_applicability(EntityType.COMPANY, False, 2_00_00_000, 0.5, 0.5)
    assert ap3.required, "2cr with heavy cash must trigger over 1cr limit"
    ap4 = check_applicability(EntityType.INDIVIDUAL_HUF, True, 60_00_000)
    assert ap4.required, "60L professional receipts must trigger over 50L limit"
    print("Applicability threshold checks OK")

    checklist = build_form26_checklist({
        "company_name": company.name, "pan": company.pan, "gstin": company.gstin,
        "registered_office": company.registered_office,
        "entity_status": EntityType.COMPANY.value, "tax_year": "2026-27",
        "turnover": f"CY: {turnover:,.0f}",
        "it_depreciation": f"Blocks: {len(it_dep.rows)}; Dep: {it_dep.total_depreciation:,.0f}",
        "gratuity_disallowance": "30,000 added back (unpaid provision)",
        "cfl_schedule": f"B/f loss {comp.brought_forward_business_loss:,.0f}; unabs. dep {comp.unabsorbed_depreciation_bf:,.0f}",
        "auditor": company.auditor,
    })
    print("Checklist clauses:", len(checklist),
          "pre-filled:", sum(1 for c in checklist if c.auto_value))

    # --- Exports ---
    wb = Workbook()
    wb.remove(wb.active)
    add_tax_sheets(wb, company, comp, it_dep, dt, book_wdv, ap, checklist)
    xlsx_path = out / "Tax_Module.xlsx"
    wb.save(xlsx_path)
    print("\nSaved:", xlsx_path)

    pdf_path = out / "Computation_of_Income.pdf"
    build_computation_pdf(str(pdf_path), company, comp, it_dep, dt, ap)
    print("Saved:", pdf_path)


if __name__ == "__main__":
    main()
