"""
Tax configuration for Tax Year 2026-27 under the Income-tax Act, 2025
(effective 1 April 2026, replacing the Income-tax Act, 1961).

EVERYTHING RATE-LIKE LIVES IN THIS ONE FILE so a practitioner can update it
when a Finance Act changes numbers, without touching engine logic.

IMPORTANT VERIFICATION NOTE (read before filing):
- The 2025 Act carried forward the 1961 Act's rate structure for companies
  and firms (per CBDT FAQs and Budget commentary). Section numbers cited
  below use the 2025 Act numbering where publicly confirmed (tax audit =
  s.63; return filing = s.263; TDS-default disallowance = s.35(b); general
  business deduction disallowances = s.37; audit-report fee = s.446).
- Budget 2026 proposals (e.g., MAT recast as a 14% "final tax", MAT credit
  set-off capped at 1/4th of new-regime liability) should be re-verified
  against the enacted Finance Act, 2026 text before reliance.
- Marginal relief on surcharge is implemented in the engine; confirm edge
  cases with the Act for amounts very close to thresholds.
"""
from __future__ import annotations

from enum import Enum


class EntityType(str, Enum):
    COMPANY = "Private/Public Limited Company (Domestic)"
    FIRM_LLP = "Partnership Firm / LLP"
    INDIVIDUAL_HUF = "Individual / HUF (Proprietorship)"


class TaxRegime(str, Enum):
    # Companies
    COMPANY_NORMAL = "Company - Normal (30% / 25% by turnover)"
    COMPANY_CONCESSIONAL_22 = "Company - Concessional 22% (old 115BAA route)"
    COMPANY_NEW_MFG_15 = "Company - New Manufacturing 15% (old 115BAB route)"
    # Firms / LLPs
    FIRM_STANDARD = "Firm/LLP - 30% flat"
    # Individuals / HUF
    IND_NEW_REGIME = "Individual/HUF - New Regime (default)"
    IND_OLD_REGIME = "Individual/HUF - Old Regime (with deductions)"


REGIMES_BY_ENTITY: dict[EntityType, list[TaxRegime]] = {
    EntityType.COMPANY: [
        TaxRegime.COMPANY_CONCESSIONAL_22,
        TaxRegime.COMPANY_NORMAL,
        TaxRegime.COMPANY_NEW_MFG_15,
    ],
    EntityType.FIRM_LLP: [TaxRegime.FIRM_STANDARD],
    EntityType.INDIVIDUAL_HUF: [TaxRegime.IND_NEW_REGIME, TaxRegime.IND_OLD_REGIME],
}


# ---------------------------------------------------------------------------
# Corporate / firm flat rates
# ---------------------------------------------------------------------------
COMPANY_NORMAL_RATE_HIGH = 0.30       # turnover above threshold in the reference year
COMPANY_NORMAL_RATE_LOW = 0.25        # turnover up to threshold in the reference year
COMPANY_TURNOVER_THRESHOLD_FOR_25 = 400_00_00_000  # Rs. 400 crore

COMPANY_CONCESSIONAL_RATE = 0.22      # old 115BAA route (no exemptions/incentives)
COMPANY_NEW_MFG_RATE = 0.15           # old 115BAB route (new manufacturing cos.)

FIRM_RATE = 0.30

# Surcharge
COMPANY_SURCHARGE_SLABS = [           # (income_above, surcharge_rate)
    (10_00_00_000, 0.12),             # > Rs. 10 crore -> 12%
    (1_00_00_000, 0.07),              # > Rs. 1 crore  -> 7%
]
COMPANY_CONCESSIONAL_SURCHARGE = 0.10  # flat 10% under 22%/15% routes
FIRM_SURCHARGE_SLABS = [
    (1_00_00_000, 0.12),              # > Rs. 1 crore -> 12%
]
HEALTH_EDUCATION_CESS = 0.04

# MAT-equivalent ("final tax" per Budget 2026 recast) -- applies to companies
# NOT under the concessional 22%/15% routes. VERIFY against Finance Act, 2026.
MAT_RATE = 0.14
MAT_APPLIES_TO = {TaxRegime.COMPANY_NORMAL}


# ---------------------------------------------------------------------------
# Individual / HUF slabs (business income; no salary standard deduction here)
# ---------------------------------------------------------------------------
# New regime (default) -- Tax Year 2026-27 slabs
IND_NEW_REGIME_SLABS = [              # (upper_limit, rate); None = no cap
    (4_00_000, 0.00),
    (8_00_000, 0.05),
    (12_00_000, 0.10),
    (16_00_000, 0.15),
    (20_00_000, 0.20),
    (24_00_000, 0.25),
    (None, 0.30),
]
IND_NEW_REGIME_REBATE_LIMIT = 12_00_000   # full rebate if total income <= this
IND_NEW_REGIME_REBATE_MAX = 60_000

# Old regime (below 60 years)
IND_OLD_REGIME_SLABS = [
    (2_50_000, 0.00),
    (5_00_000, 0.05),
    (10_00_000, 0.20),
    (None, 0.30),
]
IND_OLD_REGIME_REBATE_LIMIT = 5_00_000
IND_OLD_REGIME_REBATE_MAX = 12_500

IND_SURCHARGE_SLABS = [               # (income_above, rate)
    (5_00_00_000, 0.37),              # capped at 0.25 in new regime (handled in engine)
    (2_00_00_000, 0.25),
    (1_00_00_000, 0.15),
    (50_00_000, 0.10),
]
IND_NEW_REGIME_SURCHARGE_CAP = 0.25


# ---------------------------------------------------------------------------
# Income-tax depreciation blocks (WDV method). Rates carried over unchanged.
# Additions put to use < 180 days get half the rate for that year.
# ---------------------------------------------------------------------------
IT_DEPRECIATION_BLOCKS: dict[str, float] = {
    "Building - Residential": 0.05,
    "Building - Office/Factory (General)": 0.10,
    "Building - Temporary Structures": 0.40,
    "Furniture and Fittings": 0.10,
    "Plant and Machinery - General": 0.15,
    "Motor Cars (other than hire)": 0.15,
    "Motor Vehicles - Commercial/Hire": 0.30,
    "Computers and Software": 0.40,
    "Intangible Assets (know-how, patents, etc.)": 0.25,
    "Ships": 0.20,
}

# Keyword hints used to auto-seed blocks from Trial Balance ledger names
BLOCK_KEYWORDS: list[tuple[list[str], str]] = [
    (["computer", "laptop", "software", "server", "printer"], "Computers and Software"),
    (["furniture", "fixture", "fitting"], "Furniture and Fittings"),
    (["motor car", "car ", "vehicle"], "Motor Cars (other than hire)"),
    (["building", "premises", "factory shed"], "Building - Office/Factory (General)"),
    (["plant", "machinery", "equipment"], "Plant and Machinery - General"),
    (["patent", "trademark", "goodwill", "license", "licence", "know-how", "intangible"],
     "Intangible Assets (know-how, patents, etc.)"),
]


# ---------------------------------------------------------------------------
# Interest (old 234A/234B/234C -> consolidated interest provisions of the
# 2025 Act; monthly rates unchanged at 1% simple per month or part thereof).
# ---------------------------------------------------------------------------
INTEREST_RATE_PER_MONTH = 0.01
ADVANCE_TAX_LIABILITY_THRESHOLD = 10_000   # advance tax obligations start here
ADVANCE_TAX_INSTALLMENTS = [               # (due-by label, cumulative %)
    ("15 June", 0.15),
    ("15 September", 0.45),
    ("15 December", 0.75),
    ("15 March", 1.00),
]


# ---------------------------------------------------------------------------
# Tax audit (s.63 of the 2025 Act; report in Form 26 under Rule 47 of the
# Income-tax Rules, 2026)
# ---------------------------------------------------------------------------
TAX_AUDIT_TURNOVER_LIMIT_BUSINESS = 1_00_00_000        # Rs. 1 crore
TAX_AUDIT_TURNOVER_LIMIT_DIGITAL = 10_00_00_000        # Rs. 10 crore (>=95% digital)
TAX_AUDIT_DIGITAL_CASH_MAX_PCT = 0.05
TAX_AUDIT_RECEIPTS_LIMIT_PROFESSION = 50_00_000        # Rs. 50 lakh
TAX_AUDIT_FEE_PCT = 0.005                              # s.446: 0.5% of turnover ...
TAX_AUDIT_FEE_CAP = 1_50_000                           # ... capped at Rs. 1.5 lakh
