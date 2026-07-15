"""
Form No. 26 clause registry -- Audit Report and Statement of Particulars
under s.63 of the Income-tax Act, 2025, read with Rule 47 of the Income-tax
Rules, 2026. Replaces Forms 3CA / 3CB / 3CD from Tax Year 2026-27.

STRUCTURE (as notified/reported):
    Part A -- Basic particulars of the assessee
    Part B -- Statement of particulars (the old 3CD, expanded to ~53 clauses,
              trigger-based Yes/No answers with mandatory schedules on "Yes")
    Part C -- Audit report where accounts are audited under another law
              (replaces Form 3CA)
    Part D -- Audit report where accounts are NOT audited under another law
              (replaces Form 3CB)

CONFIRMED ANCHORS from the Rules/official FAQs as publicly reported:
    - Clauses 49-51: TDS/TCS reporting (replacing old 3CD Clause 34), incl.
      transaction counts/values after latest correction statement, interest
      under s.398(3)(a), and disallowance u/s 35(b) for TDS defaults.
    - Depreciation reported by usage-period class (>180 days / <=180 days).
    - Audit observations must be categorized (test-check / management
      representation / unable to verify) with net P&L impact.
    - UDIN mandatory (s.515(3)(b)); fee for default u/s 446.

MAINTENANCE NOTE: clause numbering between the confirmed anchors is drawn
from the Draft Rules commentary and mapped from the corresponding 3CD
clauses; where CBDT's final utility differs, edit ONLY this file -- the
engine and UI render whatever is registered here.

Each clause: (number, title, guidance, old_3cd_ref, auto_key)
`auto_key` lets the engine pre-fill answers it can compute from the TB /
computation (None = purely manual).
"""
from __future__ import annotations

FORM26_PART_B_CLAUSES: list[dict] = [
    # --- Identification & general (Part B, opening block) ---
    dict(no=1,  title="Name of the assessee", old="3CD-1", auto="company_name"),
    dict(no=2,  title="Address", old="3CD-2", auto="registered_office"),
    dict(no=3,  title="PAN", old="3CD-3", auto="pan"),
    dict(no=4,  title="Registration under indirect tax laws (GST etc.) - registration numbers",
         old="3CD-4", auto="gstin"),
    dict(no=5,  title="Status (company/firm/individual etc.)", old="3CD-5", auto="entity_status"),
    dict(no=6,  title="Tax year covered (period of audit)", old="3CD-6/8", auto="tax_year"),
    dict(no=7,  title="Whether assessed under presumptive scheme; if so, section and reason for audit",
         old="3CD-8/12", auto=None),
    dict(no=8,  title="Nature of business or profession, principal place of business; changes during the year",
         old="3CD-10", auto=None,
         guidance="Any change in nature of business must be reported with details."),
    dict(no=9,  title="Partners/members and profit-sharing ratios; changes during the year (firms/LLPs/AOPs)",
         old="3CD-9", auto=None),
    dict(no=10, title="Books of account prescribed, maintained, and examined (incl. whether computerised, location)",
         old="3CD-11", auto=None,
         guidance="Under Rule 46 (books of account) -- list cash book, journal, ledger etc."),
    dict(no=11, title="Method of accounting employed; change in method from immediately preceding tax year",
         old="3CD-13(a)-(c)", auto=None),
    dict(no=12, title="Adjustments required by Income Computation and Disclosure Standards (ICDS I-X) with schedule",
         old="3CD-13(d)-(f)", auto=None,
         guidance="ICDS-wise impact statement is a structured schedule in Form 26."),
    dict(no=13, title="Method of valuation of closing stock; deviation from prescribed method and effect on profit",
         old="3CD-14", auto=None),
    dict(no=14, title="Capital asset converted into stock-in-trade -- particulars",
         old="3CD-15", auto=None),
    # --- Income recognition block ---
    dict(no=15, title="Amounts not credited to P&L but chargeable to tax (items of income, capital receipts etc.)",
         old="3CD-16", auto=None,
         guidance="Expanded list incl. buybacks, subsidies, business-trust income."),
    dict(no=16, title="Receipts/transactions of a capital nature; deemed income items",
         old="3CD-16 (part)", auto=None),
    dict(no=17, title="Transfer of land/building below stamp duty value (old 43CA/50C pattern) -- details",
         old="3CD-17", auto=None),
    # --- Depreciation & deductions block ---
    dict(no=18, title="Depreciation allowable: block-wise WDV, rate, additions by usage period (>180 / <=180 days), deletions",
         old="3CD-18", auto="it_depreciation",
         guidance="Form 26 requires usage-period classification only (no asset-wise put-to-use dates)."),
    dict(no=19, title="Amounts admissible under incentive deduction provisions (old 33AB/35 family)",
         old="3CD-19", auto=None),
    dict(no=20, title="Bonus/commission to employees; sums received from employees (PF/ESI) with due-date compliance",
         old="3CD-20", auto=None,
         guidance="ESI/PF reporting limited to disallowable amounts under Form 26."),
    # --- Disallowances block ---
    dict(no=21, title="Amounts debited to P&L being capital, personal, advertisement in political publications etc.",
         old="3CD-21(a)", auto=None),
    dict(no=22, title="Expenditure on entertainment/club/penalty or offence-linked payments",
         old="3CD-21 (parts)", auto=None,
         guidance="Penalties and interest on TDS/late fees are disallowable u/s 37 pattern."),
    dict(no=23, title="Disallowance for TDS/TCS default on payments (old 40(a) pattern -> s.35(b), IT Act 2025)",
         old="3CD-21(b)", auto="tds_default_disallowance"),
    dict(no=24, title="Payments to specified/related persons -- excessive or unreasonable (old 40A(2)(b))",
         old="3CD-23", auto=None),
    dict(no=25, title="Cash payments above threshold (old 40A(3)/(3A)) -- schedule of violations",
         old="3CD-21(d)", auto=None),
    dict(no=26, title="Provision for gratuity not allowable (unapproved fund / unpaid)",
         old="3CD-21(e)", auto="gratuity_disallowance"),
    dict(no=27, title="Sums payable deductible only on actual payment (old 43B) -- paid/unpaid by due date schedule",
         old="3CD-26", auto=None),
    dict(no=28, title="Amounts of interest inadmissible (MSME payees -- old 23 MSMED interplay; s.35 family)",
         old="3CD-22", auto=None),
    dict(no=29, title="Deemed profits and gains (old 41 pattern -- remission/cessation of liability)",
         old="3CD-25", auto=None),
    # --- Loans, deposits, receipts block ---
    dict(no=30, title="Acceptance of loans/deposits above threshold otherwise than by account-payee mode (old 269SS)",
         old="3CD-31(a)", auto=None),
    dict(no=31, title="Receipt of Rs. 2 lakh+ otherwise than by prescribed modes (old 269ST)",
         old="3CD-31(ba)", auto=None),
    dict(no=32, title="Repayment of loans/deposits above threshold otherwise than by account-payee mode (old 269T)",
         old="3CD-31(c)", auto=None),
    # --- Losses & carry-forwards block ---
    dict(no=33, title="Details of brought forward loss / unabsorbed depreciation with origin year and set-off",
         old="3CD-32", auto="cfl_schedule",
         guidance="Origin-year-wise carry-forward table; matches the CFL schedule."),
    dict(no=34, title="Change in shareholding affecting carry-forward of losses (old 79 pattern)",
         old="3CD-32(b)", auto=None),
    dict(no=35, title="Speculation loss / specified-business loss details",
         old="3CD-32(c)-(e)", auto=None),
    # --- Chapter deductions & exempt income ---
    dict(no=36, title="Deductions admissible under the deduction chapter (old Ch. VI-A -> ss.123-140, IT Act 2025)",
         old="3CD-33", auto=None),
    dict(no=37, title="Exempt income and related expenditure disallowance (old 14A pattern)",
         old="3CD (14A rider)", auto=None),
    # --- Audit & compliance status block ---
    dict(no=38, title="Whether liable to audit under any other law; details of that audit",
         old="3CA context", auto=None),
    dict(no=39, title="Cost audit / excise audit / other special audits carried out -- details",
         old="3CD-37/38/39", auto=None),
    dict(no=40, title="Quantitative details of principal items of goods traded/manufactured (opening, purchases, sales, closing)",
         old="3CD-35", auto=None,
         guidance="Stock reconciliation discipline revived in Form 26."),
    dict(no=41, title="Accounting ratios with previous-year comparison (GP%, NP%, stock-to-turnover etc.)",
         old="3CD-40", auto="ratios"),
    dict(no=42, title="Total turnover / gross receipts of the tax year and the immediately preceding tax year",
         old="3CD-40 (part)", auto="turnover"),
    # --- Transactions & surveillance block ---
    dict(no=43, title="Transactions with related parties / specified domestic transactions summary",
         old="3CD-23/92E context", auto=None),
    dict(no=44, title="GST turnover reconciliation with books (GSTR vs P&L) -- summary level",
         old="3CD-44 (recast)", auto=None,
         guidance="Form 26 asks summary-level GST linkage, not invoice-level."),
    dict(no=45, title="Break-up of expenditure: registered vs unregistered GST suppliers",
         old="3CD-44", auto=None),
    dict(no=46, title="Details of demands raised / refunds issued under other tax laws during the year",
         old="3CD-41", auto=None),
    dict(no=47, title="Dividend distributions incl. deemed dividend u/s 2(22)(f)-pattern reporting",
         old="3CD-36A/36B", auto=None),
    dict(no=48, title="IT systems / digital books environment: accounting software, audit trail, data location",
         old="(new)", auto=None,
         guidance="New digital-governance block: software used, audit-trail status, server/data location."),
    # --- TDS/TCS block (CONFIRMED clause anchors) ---
    dict(no=49, title="TDS/TCS: section-wise deduction/collection compliance schedule",
         old="3CD-34(a)", auto=None,
         guidance="Section references use the consolidated s.393 family of the 2025 Act."),
    dict(no=50, title="TDS/TCS statements: total transactions per latest correction statement, transactions NOT reported (count and value)",
         old="3CD-34(b)", auto=None,
         guidance="Quantification replaces the old Yes/No -- pull TRACES post-correction data."),
    dict(no=51, title="Interest under s.398(3)(a) for late deduction / non-deduction / late deposit; disallowance u/s 35(b) schedule",
         old="3CD-34(c)", auto=None),
    # --- Closing block ---
    dict(no=52, title="Audit observations/qualifications categorized (test-check / management representation / unable to verify) with net P&L impact",
         old="3CA/3CB para 3", auto=None,
         guidance="Mandatory three-way categorization is new in Form 26."),
    dict(no=53, title="Accountant's particulars: name, membership no., FRN, UDIN (s.515(3)(b)), place, date, signature",
         old="3CD-close", auto="auditor"),
]
