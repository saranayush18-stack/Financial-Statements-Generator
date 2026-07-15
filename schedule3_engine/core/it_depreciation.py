"""
Income-tax depreciation (block-of-assets, WDV method).

Mirrors the Book1 'IT Depreciation' sheet: per block --
    rate | opening WDV | additions >=180 days | additions <180 days |
    deletions | total | depreciation | closing WDV
Depreciation = (opening + additions_180plus - deletions) * rate
             + additions_less180 * rate / 2
floored so a block never depreciates below zero.

Every row is editable/overridable by the user before the computation is
finalized -- the auto-seed from the Trial Balance is only a starting point,
because a TB alone cannot know put-to-use dates.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from data.tax_config import IT_DEPRECIATION_BLOCKS, BLOCK_KEYWORDS
from models import TrialBalance, MappingEntry


@dataclass
class AssetBlockRow:
    block_name: str
    rate: float
    opening_wdv: float = 0.0
    additions_180_plus: float = 0.0
    additions_less_180: float = 0.0
    deletions: float = 0.0

    @property
    def total(self) -> float:
        return round(self.opening_wdv + self.additions_180_plus
                     + self.additions_less_180 - self.deletions, 2)

    @property
    def depreciation(self) -> float:
        base = max(self.opening_wdv + self.additions_180_plus - self.deletions, 0.0)
        dep = base * self.rate + max(self.additions_less_180, 0.0) * self.rate / 2
        return round(min(dep, max(self.total, 0.0)), 2)

    @property
    def closing_wdv(self) -> float:
        return round(self.total - self.depreciation, 2)


@dataclass
class ITDepreciationSchedule:
    rows: list[AssetBlockRow] = field(default_factory=list)

    @property
    def total_opening(self) -> float:
        return round(sum(r.opening_wdv for r in self.rows), 2)

    @property
    def total_additions_180_plus(self) -> float:
        return round(sum(r.additions_180_plus for r in self.rows), 2)

    @property
    def total_additions_less_180(self) -> float:
        return round(sum(r.additions_less_180 for r in self.rows), 2)

    @property
    def total_deletions(self) -> float:
        return round(sum(r.deletions for r in self.rows), 2)

    @property
    def total_depreciation(self) -> float:
        return round(sum(r.depreciation for r in self.rows), 2)

    @property
    def total_closing_wdv(self) -> float:
        return round(sum(r.closing_wdv for r in self.rows), 2)


def _match_block(ledger_name: str) -> str | None:
    low = ledger_name.lower()
    for keywords, block in BLOCK_KEYWORDS:
        if any(k in low for k in keywords):
            return block
    return None


def seed_schedule_from_tb(tb: TrialBalance, mappings: dict[str, MappingEntry]) -> ITDepreciationSchedule:
    """Best-effort starting point: gross fixed-asset ledgers grouped into IT
    blocks using keyword hints, previous-year closing treated as opening WDV
    and the current-year increase treated as an addition put to use >=180
    days (the user corrects put-to-use split and deletions in the UI).
    Accumulated-depreciation ledgers are ignored -- IT WDV is not book WDV."""
    acc: dict[str, dict[str, float]] = {}
    for ledger in tb.ledgers:
        m = mappings.get(ledger.ledger_name)
        if m is None or m.sub_head not in (
            "Property, Plant and Equipment", "Intangible Assets", "Capital Work-in-Progress",
        ):
            continue
        if "accumulated" in ledger.ledger_name.lower() or "depreciation" in ledger.ledger_name.lower():
            continue
        if m.sub_head == "Capital Work-in-Progress":
            continue  # CWIP is not depreciable until put to use
        block = _match_block(ledger.ledger_name) or (
            "Intangible Assets (know-how, patents, etc.)" if m.sub_head == "Intangible Assets"
            else "Plant and Machinery - General")
        slot = acc.setdefault(block, {"opening": 0.0, "additions": 0.0})
        py = ledger.previous_year_closing or 0.0
        cy = ledger.closing_balance
        slot["opening"] += py
        slot["additions"] += max(cy - py, 0.0)

    schedule = ITDepreciationSchedule()
    for block, vals in acc.items():
        schedule.rows.append(AssetBlockRow(
            block_name=block,
            rate=IT_DEPRECIATION_BLOCKS.get(block, 0.15),
            opening_wdv=round(vals["opening"], 2),
            additions_180_plus=round(vals["additions"], 2),
        ))
    schedule.rows.sort(key=lambda r: r.block_name)
    return schedule
