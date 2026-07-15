"""
Headless smoke test of the Streamlit app's RESULTS stage — the stage that
contains the entire Income Tax module (computation, IT depreciation editor,
deferred tax, tax audit checklist, tax downloads).

Approach: build tb / mappings / company exactly the way demo.py does, inject
them into session_state via streamlit.testing.v1.AppTest, set stage="results",
run, and assert no exceptions were raised anywhere in the script.
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from streamlit.testing.v1 import AppTest

from models import Company
from core.tb_parser import parse_trial_balance
from core.mapping_store import MappingStore, resolve_mapping
from sample_data.generate_sample_tb import generate as generate_sample_tb


def build_state(tmpdir: str):
    tb_path = generate_sample_tb(str(Path(tmpdir) / "sample_trial_balance.xlsx"))
    company = Company(
        name="Bharat Precision Components Private Limited",
        cin="U29100MH2015PTC123456",
        pan="AABCB1234C",
        gstin="27AABCB1234C1Z5",
        registered_office="Plot No. 45, MIDC Industrial Area, Pune, Maharashtra - 411019",
        auditor="M/s ABC & Associates, Chartered Accountants",
        directors=["Mr. Rajesh Sharma", "Mrs. Priya Sharma"],
        financial_year_start=date(2025, 4, 1),
        financial_year_end=date(2026, 3, 31),
        company_id=1,
    )
    tb, parse_warnings = parse_trial_balance(
        tb_path, company, "FY 2025-26 (1 April 2025 to 31 March 2026)")
    store = MappingStore(str(Path(tmpdir) / "mappings.db"))
    mappings, unmapped = {}, []
    for ledger in tb.ledgers:
        m = resolve_mapping(store, company.company_id, ledger.ledger_name)
        if m is None:
            unmapped.append(ledger.ledger_name)
        else:
            mappings[ledger.ledger_name] = m
    assert not unmapped, f"sample TB has unmapped ledgers: {unmapped}"
    return tb, mappings, company


def main():
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
        tb, mappings, company = build_state(tmpdir)

        at = AppTest.from_file("streamlit_app.py", default_timeout=300)
        at.session_state["stage"] = "results"
        at.session_state["tb"] = tb
        at.session_state["mappings"] = mappings
        at.session_state["company"] = company
        at.session_state["fy_label"] = "FY 2025-26 (1 April 2025 to 31 March 2026)"
        at.run()

        if at.exception:
            for exc in at.exception:
                print("EXCEPTION:", exc.value)
                print(exc.stack_trace)
            raise SystemExit(1)

        # Structural assertions: the tax section actually rendered
        subheaders = [h.value for h in at.subheader]
        assert any("Income Tax" in s for s in subheaders), subheaders
        tab_labels = [t.label for t in at.tabs]
        for expected in ("Income Computation", "IT Depreciation", "Deferred Tax",
                         "Tax Audit (Form 26)"):
            assert expected in tab_labels, (expected, tab_labels)

        metrics = {m.label: m.value for m in at.metric}
        for expected in ("Taxable Income", "Total Tax Liability", "Net Payable / (Refund)",
                         "Total IT Depreciation", "Closing DTA/(DTL)"):
            assert expected in metrics, (expected, list(metrics))
        print("Metrics:", {k: metrics[k] for k in (
            "Taxable Income", "Total Tax Liability", "Net Payable / (Refund)",
            "Total IT Depreciation", "Closing DTA/(DTL)")})

        # ---- Interaction 1: switch entity type to Firm/LLP (regime list must follow)
        at.selectbox[0].select("Partnership Firm / LLP").run()
        assert not at.exception, [e.value for e in at.exception]

        # ---- Interaction 2: back to company, pick concessional 22% regime
        at.selectbox[0].select("Private/Public Limited Company (Domestic)").run()
        regime_options = at.selectbox[1].options
        pick = next(o for o in regime_options if "22" in o)
        at.selectbox[1].select(pick).run()
        assert not at.exception, [e.value for e in at.exception]

        # ---- Interaction 3: override one adjustment line via its checkbox
        ov = next(cb for cb in at.checkbox if cb.key and cb.key.startswith("ov_"))
        ov.check().run()
        assert not at.exception, [e.value for e in at.exception]

        # ---- Interaction 4: type a clause response in the audit checklist
        cl = next(ta for ta in at.text_area if ta.key and ta.key.startswith("cl_"))
        cl.input("Verified from books; no adverse remarks.").run()
        assert not at.exception, [e.value for e in at.exception]

        print("STREAMLIT APPTEST OK — results stage + tax module render and interact cleanly")


if __name__ == "__main__":
    main()
