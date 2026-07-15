"""
Tax Audit under s.63, Income-tax Act, 2025 (report in Form No. 26, Rule 47
of the Income-tax Rules, 2026).

Two responsibilities:
1.  APPLICABILITY: given turnover / gross receipts / cash percentages /
    presumptive-scheme facts, decide whether audit is required, with the
    reasoning spelled out (so the app can show *why*).
2.  CHECKLIST: materialize the Form 26 Part B clause registry into a
    working checklist, auto-filling what the engine already knows
    (turnover, depreciation schedule, ratios, carry-forward losses),
    leaving the rest for the auditor with status tracking.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from data.tax_config import (
    EntityType,
    TAX_AUDIT_TURNOVER_LIMIT_BUSINESS, TAX_AUDIT_TURNOVER_LIMIT_DIGITAL,
    TAX_AUDIT_DIGITAL_CASH_MAX_PCT, TAX_AUDIT_RECEIPTS_LIMIT_PROFESSION,
    TAX_AUDIT_FEE_PCT, TAX_AUDIT_FEE_CAP,
)
from data.form26_clauses import FORM26_PART_B_CLAUSES


@dataclass
class AuditApplicability:
    required: bool
    reasons: list[str] = field(default_factory=list)
    applicable_limit: float | None = None
    fee_if_defaulted: float = 0.0
    form: str = "Form No. 26 (Rule 47, Income-tax Rules, 2026)"
    due_note: str = ("File one month before the return due date under "
                     "s.263(1), Income-tax Act, 2025; UDIN mandatory.")


def check_applicability(
    entity_type: EntityType,
    is_profession: bool,
    turnover_or_receipts: float,
    cash_receipts_pct: float = 1.0,
    cash_payments_pct: float = 1.0,
    presumptive_scheme_opted: bool = False,
    declared_below_presumptive: bool = False,
) -> AuditApplicability:
    reasons: list[str] = []
    required = False
    limit: float | None = None

    if is_profession:
        limit = TAX_AUDIT_RECEIPTS_LIMIT_PROFESSION
        if turnover_or_receipts > limit:
            required = True
            reasons.append(
                f"Gross professional receipts Rs. {turnover_or_receipts:,.0f} exceed the "
                f"Rs. {limit:,.0f} limit for professions (s.63)."
            )
        else:
            reasons.append(
                f"Gross professional receipts Rs. {turnover_or_receipts:,.0f} are within the "
                f"Rs. {limit:,.0f} profession limit."
            )
    else:
        digital_ok = (cash_receipts_pct <= TAX_AUDIT_DIGITAL_CASH_MAX_PCT
                      and cash_payments_pct <= TAX_AUDIT_DIGITAL_CASH_MAX_PCT)
        limit = TAX_AUDIT_TURNOVER_LIMIT_DIGITAL if digital_ok else TAX_AUDIT_TURNOVER_LIMIT_BUSINESS
        if digital_ok:
            reasons.append(
                "Cash receipts and cash payments are each within 5% of totals, so the "
                f"enhanced Rs. {TAX_AUDIT_TURNOVER_LIMIT_DIGITAL:,.0f} limit applies."
            )
        else:
            reasons.append(
                "Cash receipts/payments exceed 5% of totals, so the standard "
                f"Rs. {TAX_AUDIT_TURNOVER_LIMIT_BUSINESS:,.0f} limit applies."
            )
        if turnover_or_receipts > limit:
            required = True
            reasons.append(
                f"Business turnover Rs. {turnover_or_receipts:,.0f} exceeds the applicable "
                f"limit of Rs. {limit:,.0f} (s.63)."
            )
        else:
            reasons.append(
                f"Business turnover Rs. {turnover_or_receipts:,.0f} is within the applicable "
                f"limit of Rs. {limit:,.0f}."
            )

    if presumptive_scheme_opted and declared_below_presumptive:
        required = True
        reasons.append(
            "Assessee opted for a presumptive scheme but declares profit below the deemed "
            "rate -- audit is mandatory irrespective of turnover."
        )

    fee = min(turnover_or_receipts * TAX_AUDIT_FEE_PCT, TAX_AUDIT_FEE_CAP) if required else 0.0
    return AuditApplicability(required=required, reasons=reasons,
                               applicable_limit=limit, fee_if_defaulted=round(fee, 2))


@dataclass
class ClauseItem:
    no: int
    title: str
    old_3cd_ref: str
    guidance: str = ""
    auto_value: str = ""          # engine-prefilled content, if any
    response: str = ""            # auditor's answer / particulars
    status: str = "Pending"       # Pending | In Progress | Completed | N/A


def build_form26_checklist(context: dict) -> list[ClauseItem]:
    """`context` carries whatever the engine knows keyed by auto_key
    (company_name, pan, turnover, it_depreciation summary, ratios, ...)."""
    items: list[ClauseItem] = []
    for c in FORM26_PART_B_CLAUSES:
        auto_val = ""
        key = c.get("auto")
        if key and key in context and context[key] not in (None, ""):
            auto_val = str(context[key])
        items.append(ClauseItem(
            no=c["no"], title=c["title"], old_3cd_ref=c.get("old", ""),
            guidance=c.get("guidance", ""), auto_value=auto_val,
            status="Completed" if auto_val else "Pending",
        ))
    return items
