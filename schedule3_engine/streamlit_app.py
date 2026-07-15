"""
AI Financial Statement Generator -- Streamlit App

Same deployment pattern as AuditLens: upload a Trial Balance, get back
Schedule III financial statements (Excel + PDF), no local installs
required once deployed to Streamlit Community Cloud.

RUN LOCALLY:
    pip install streamlit pandas openpyxl reportlab
    streamlit run streamlit_app.py

DEPLOY TO STREAMLIT COMMUNITY CLOUD (same steps as AuditLens):
    1. Push this whole folder to a GitHub repo.
    2. Go to share.streamlit.io -> "New app" -> pick the repo.
    3. Set "Main file path" to streamlit_app.py.
    4. Deploy.
"""
from __future__ import annotations

import io
from datetime import date, datetime
from pathlib import Path

import streamlit as st

from models import Company, MappingEntry, Statement, CurrentNonCurrent, Nature
from core.tb_parser import parse_trial_balance
from core.mapping_store import MappingStore, resolve_mapping
from core.validator import validate, summarize_issues
from core.statement_generator import (
    generate_balance_sheet, generate_profit_and_loss, generate_cash_flow_indirect,
    carry_profit_into_reserves,
)
from core.notes_generator import generate_notes
from core.ratios import compute_ratios
from core.soce_generator import generate_soce, soce_reconciles_to_balance_sheet
from core.ageing import (
    parse_ageing_file, build_ageing_grid, unavailable_grid, ageing_validation_issues,
)
from export.excel_export import build_workbook, save_workbook
from export.pdf_export import build_pdf
from data.classification_rules import SCHEDULE_III_STRUCTURE, RULES
from core.it_depreciation import seed_schedule_from_tb, AssetBlockRow, ITDepreciationSchedule
from core.income_computation import build_standard_computation
from core.deferred_tax import build_standard_deferred_tax, effective_tax_rate
from core.tax_audit import check_applicability, build_form26_checklist
from data.tax_config import (
    EntityType, TaxRegime, REGIMES_BY_ENTITY, IT_DEPRECIATION_BLOCKS,
    COMPANY_CONCESSIONAL_SURCHARGE,
)
from export.tax_export import add_tax_sheets, build_computation_pdf

st.set_page_config(page_title="AI Financial Statement Generator", page_icon="📊", layout="wide")

MAPPING_DB_PATH = "schedule3_mappings.db"  # persists learned mappings across sessions on the same deployment


# ---------------------------------------------------------------------------
# Build a lookup of sub_head -> (major_head, statement, current/non-current,
# nature, note_ref) from the rule table, for the manual-mapping dropdown.
# ---------------------------------------------------------------------------
@st.cache_data
def subhead_metadata() -> dict[str, dict]:
    meta = {}

    # BALANCE_SHEET is 3 levels deep: section (Equity and Liabilities / Assets)
    # -> major_head -> [sub_heads]. PROFIT_AND_LOSS is 2 levels: major_head -> [sub_heads].
    bs_struct = SCHEDULE_III_STRUCTURE["BALANCE_SHEET"]
    for section, majors in bs_struct.items():
        for major_head, subs in majors.items():
            for sub in subs:
                meta.setdefault(sub, {"major_head": major_head, "statement": Statement.BALANCE_SHEET})

    pnl_struct = SCHEDULE_III_STRUCTURE["PROFIT_AND_LOSS"]
    for major_head, subs in pnl_struct.items():
        for sub in subs:
            meta.setdefault(sub, {"major_head": major_head, "statement": Statement.PROFIT_AND_LOSS})

    # fill nature / current_or_non_current / note_ref from RULES (first match wins)
    for keywords, major_head, sub_head, statement, cur_ncur, nature, note_ref in RULES:
        if sub_head in meta:
            meta[sub_head].setdefault("current_or_non_current", cur_ncur)
            meta[sub_head].setdefault("nature", nature)
            meta[sub_head].setdefault("note_ref", note_ref)

    # A handful of sub_heads exist in the taxonomy (as valid Schedule III
    # categories) but have no ledger keyword rule pointing at them yet --
    # meaning the automatic classifier can never produce them, but a user
    # could still pick one manually. Without this, the nature/current-vs-
    # non-current would silently default to DEBIT/NOT_APPLICABLE, which is
    # wrong for several of these (e.g. Deferred Tax Liabilities is a
    # CREDIT-nature, non-current item). Fill them in explicitly.
    FALLBACK_META = {
        "Deferred Tax Liabilities (Net)": (CurrentNonCurrent.NON_CURRENT, Nature.CREDIT,
                                            "Note - Deferred Tax Liabilities (Net)"),
        "Other Long-Term Liabilities": (CurrentNonCurrent.NON_CURRENT, Nature.CREDIT,
                                         "Note - Other Long-Term Liabilities"),
        "Long-Term Loans and Advances": (CurrentNonCurrent.NON_CURRENT, Nature.DEBIT,
                                          "Note - Long-Term Loans and Advances"),
        "Other Non-Current Assets": (CurrentNonCurrent.NON_CURRENT, Nature.DEBIT,
                                      "Note - Other Non-Current Assets"),
        "Current Investments": (CurrentNonCurrent.CURRENT, Nature.DEBIT,
                                 "Note - Current Investments"),
        "Changes in Inventories of Finished Goods, WIP and Stock-in-Trade": (
            CurrentNonCurrent.NOT_APPLICABLE, Nature.DEBIT,
            "Note - Changes in Inventories"),
    }
    for sub_head, (cur_ncur, nature, note_ref) in FALLBACK_META.items():
        if sub_head in meta:
            meta[sub_head].setdefault("current_or_non_current", cur_ncur)
            meta[sub_head].setdefault("nature", nature)
            meta[sub_head].setdefault("note_ref", note_ref)

    return meta


def all_sub_heads() -> list[str]:
    return sorted(subhead_metadata().keys())


# ---------------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------------
def init_state():
    defaults = {
        "stage": "upload",       # upload -> mapping -> results
        "tb": None,
        "mappings": {},
        "unmapped": [],
        "manual_choices": {},
        "company": None,
        "fy_label": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()

st.title("📊 AI Financial Statement Generator")
st.caption("Trial Balance → Schedule III Financial Statements. Rule-based classification, fully auditable, no API cost.")

# ---------------------------------------------------------------------------
# SIDEBAR -- Company details
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Company Details")
    company_name = st.text_input("Company Name*", value="")
    cin = st.text_input("CIN")
    pan = st.text_input("PAN")
    gstin = st.text_input("GSTIN")
    registered_office = st.text_area("Registered Office", height=68)
    auditor = st.text_input("Auditor", value="")
    directors_raw = st.text_area("Directors (one per line)", height=68)

    st.subheader("Financial Year")
    col1, col2 = st.columns(2)
    fy_start = col1.date_input("FY Start", value=date(date.today().year, 4, 1))
    fy_end = col2.date_input("FY End", value=date(date.today().year + 1, 3, 31))
    fy_label = st.text_input(
        "FY Label (shown on statements)",
        value=f"FY {fy_start.year}-{str(fy_end.year)[-2:]} ({fy_start.strftime('%-d %B %Y')} to {fy_end.strftime('%-d %B %Y')})",
    )

    st.divider()
    if st.button("🔄 Start Over", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ---------------------------------------------------------------------------
# STAGE 1: Upload
# ---------------------------------------------------------------------------
if st.session_state.stage == "upload":
    st.subheader("1. Upload Trial Balance")
    st.write(
        "Accepts Excel or CSV exports from Tally, Busy, Zoho, QuickBooks, or SAP. "
        "Needs at minimum: Ledger Name, Debit, Credit columns (closing balances, not movements)."
    )
    tb_file = st.file_uploader("Trial Balance", type=["xlsx", "xls", "csv"])

    with st.expander("Optional: Trade Receivables / Payables ageing (party-wise, with due dates)"):
        st.write(
            "A plain Trial Balance has no due-date detail, so the statutory ageing schedule "
            "can't be computed from it alone. If you have a debtors/creditors ageing export "
            "(Party Name, Amount, Due Date, and optionally Disputed/Doubtful columns), upload it "
            "here. If you skip this, the statements will still generate -- the ageing note will "
            "just say it isn't available yet, instead of guessing."
        )
        rec_ageing_file = st.file_uploader("Trade Receivables ageing", type=["xlsx", "xls", "csv"], key="rec_ageing")
        pay_ageing_file = st.file_uploader("Trade Payables ageing", type=["xlsx", "xls", "csv"], key="pay_ageing")

    generate_clicked = st.button("Generate Financial Statements ▶", type="primary", disabled=not (tb_file and company_name))
    if not company_name:
        st.caption("⚠️ Company Name is required (see sidebar).")

    if generate_clicked:
        with st.spinner("Parsing Trial Balance..."):
            company = Company(
                name=company_name, cin=cin or None, pan=pan or None, gstin=gstin or None,
                registered_office=registered_office or None, auditor=auditor or None,
                directors=[d.strip() for d in directors_raw.splitlines() if d.strip()],
                financial_year_start=fy_start, financial_year_end=fy_end, company_id=1,
            )
            # tb_parser expects a file path; write the upload to a temp buffer first
            suffix = Path(tb_file.name).suffix
            tmp_path = f"/tmp/_uploaded_tb{suffix}"
            with open(tmp_path, "wb") as f:
                f.write(tb_file.getbuffer())
            tb, parse_warnings = parse_trial_balance(tmp_path, company, fy_label)

        store = MappingStore(MAPPING_DB_PATH)
        mappings, unmapped = {}, []
        for ledger in tb.ledgers:
            resolved = resolve_mapping(store, company.company_id, ledger.ledger_name)
            if resolved:
                mappings[ledger.ledger_name] = resolved
                if resolved.source == "RULE_ENGINE":
                    store.save_company_mapping(company.company_id, resolved, user_name="streamlit_user")
            else:
                unmapped.append(ledger.ledger_name)

        st.session_state.tb = tb
        st.session_state.parse_warnings = parse_warnings
        st.session_state.mappings = mappings
        st.session_state.unmapped = unmapped
        st.session_state.company = company
        st.session_state.fy_label = fy_label
        st.session_state.rec_ageing_bytes = rec_ageing_file.getbuffer().tobytes() if rec_ageing_file else None
        st.session_state.rec_ageing_suffix = Path(rec_ageing_file.name).suffix if rec_ageing_file else None
        st.session_state.pay_ageing_bytes = pay_ageing_file.getbuffer().tobytes() if pay_ageing_file else None
        st.session_state.pay_ageing_suffix = Path(pay_ageing_file.name).suffix if pay_ageing_file else None

        st.session_state.stage = "mapping" if unmapped else "results"
        st.rerun()

# ---------------------------------------------------------------------------
# STAGE 2: Manual mapping for unclassified ledgers
# ---------------------------------------------------------------------------
elif st.session_state.stage == "mapping":
    st.subheader("2. Map Unclassified Ledgers")
    st.write(
        f"{len(st.session_state.unmapped)} ledger(s) didn't match the rule engine. "
        "Map them below -- your choice is saved permanently, so this ledger auto-classifies on every future upload."
    )

    options = all_sub_heads()
    with st.form("mapping_form"):
        choices = {}
        for ledger_name in st.session_state.unmapped:
            choices[ledger_name] = st.selectbox(ledger_name, options=["-- select --"] + options, key=f"map_{ledger_name}")
        submitted = st.form_submit_button("Save Mappings & Continue ▶", type="primary")

    if submitted:
        meta = subhead_metadata()
        store = MappingStore(MAPPING_DB_PATH)
        still_unmapped = []
        for ledger_name, sub_head in choices.items():
            if sub_head == "-- select --":
                still_unmapped.append(ledger_name)
                continue
            m = meta[sub_head]
            entry = MappingEntry(
                ledger_name=ledger_name, major_head=m["major_head"], sub_head=sub_head,
                statement=m["statement"], current_or_non_current=m.get("current_or_non_current", CurrentNonCurrent.NOT_APPLICABLE),
                nature=m.get("nature", Nature.DEBIT), confidence=1.0, source="MANUAL",
                note_ref=m.get("note_ref"),
            )
            st.session_state.mappings[ledger_name] = entry
            store.save_company_mapping(st.session_state.company.company_id, entry, user_name="streamlit_user")

        if still_unmapped:
            st.warning(f"{len(still_unmapped)} ledger(s) still unmapped -- please select a head for each before continuing.")
            st.session_state.unmapped = still_unmapped
        else:
            st.session_state.unmapped = []
            st.session_state.stage = "results"
        st.rerun()

    if st.button("⬅ Back to upload"):
        st.session_state.stage = "upload"
        st.rerun()

# ---------------------------------------------------------------------------
# STAGE 3: Results
# ---------------------------------------------------------------------------
elif st.session_state.stage == "results":
    st.subheader("3. Financial Statements")

    tb = st.session_state.tb
    company = st.session_state.company
    fy_label = st.session_state.fy_label
    mappings = st.session_state.mappings

    for w in st.session_state.get("parse_warnings", []):
        st.warning(w)

    issues = validate(tb, mappings)
    bs = generate_balance_sheet(tb, mappings)
    pnl = generate_profit_and_loss(tb, mappings)
    bs = carry_profit_into_reserves(bs, pnl)
    cash_flow = generate_cash_flow_indirect(tb, mappings, pnl, bs)
    soce = generate_soce(bs, pnl)
    soce_ok = soce_reconciles_to_balance_sheet(soce, bs)

    def bs_amount(sub_head_name: str) -> float:
        for section in (bs.equity_and_liabilities, bs.assets):
            for major in section:
                for sh in major.sub_heads:
                    if sh.sub_head == sub_head_name:
                        return sh.current_year
        return 0.0

    def load_ageing(kind: str, bytes_key: str, suffix_key: str, bs_sub_head: str):
        b, suf = st.session_state.get(bytes_key), st.session_state.get(suffix_key)
        if not b:
            return unavailable_grid(kind, company.financial_year_end,
                                     "no party-wise ledger with due dates was supplied.")
        tmp = f"/tmp/_{bytes_key}{suf}"
        with open(tmp, "wb") as f:
            f.write(b)
        try:
            items = parse_ageing_file(tmp)
            return build_ageing_grid(kind, items, as_of=company.financial_year_end,
                                      balance_sheet_amount=bs_amount(bs_sub_head))
        except ValueError as exc:
            return unavailable_grid(kind, company.financial_year_end, str(exc))

    receivables_ageing = load_ageing("Trade Receivables", "rec_ageing_bytes", "rec_ageing_suffix", "Trade Receivables")
    payables_ageing = load_ageing("Trade Payables", "pay_ageing_bytes", "pay_ageing_suffix", "Trade Payables")
    issues += ageing_validation_issues(receivables_ageing)
    issues += ageing_validation_issues(payables_ageing)

    notes = generate_notes(bs, pnl)
    ratios = compute_ratios(bs, pnl)

    # --- Top metrics ---
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Balance Sheet Tallies", "✅ Yes" if bs.is_tallied else "❌ No")
    m2.metric("Total Assets (CY)", f"₹{bs.total_assets_cy:,.0f}")
    m3.metric("Profit After Tax (CY)", f"₹{pnl.profit_after_tax_cy:,.0f}")
    vsum = summarize_issues(issues)
    m4.metric("Validation Issues", f"{vsum['errors']} errors, {vsum['warnings']} warnings")

    if not bs.is_tallied:
        st.error("Balance Sheet does not tally -- review mappings before issuing these statements.")
    if not soce_ok:
        st.error("Statement of Changes in Equity does not reconcile to the Balance Sheet.")

    # --- Downloads ---
    wb = build_workbook(company, fy_label, tb, mappings, bs, pnl, cash_flow, notes, issues, ratios,
                         soce=soce, receivables_ageing=receivables_ageing, payables_ageing=payables_ageing)
    excel_buf = io.BytesIO()
    wb.save(excel_buf)
    excel_buf.seek(0)

    pdf_tmp = "/tmp/_Financial_Statements.pdf"
    build_pdf(pdf_tmp, company, fy_label, bs, pnl, cash_flow, notes,
              soce=soce, receivables_ageing=receivables_ageing, payables_ageing=payables_ageing)
    with open(pdf_tmp, "rb") as f:
        pdf_bytes = f.read()

    dl1, dl2 = st.columns(2)
    dl1.download_button("⬇ Download Excel", data=excel_buf, file_name=f"{company.name}_Financial_Statements.xlsx",
                         use_container_width=True)
    dl2.download_button("⬇ Download PDF", data=pdf_bytes, file_name=f"{company.name}_Financial_Statements.pdf",
                         use_container_width=True)

    # --- Tabs for on-screen review ---
    tab_bs, tab_pnl, tab_cf, tab_soce, tab_ageing, tab_val = st.tabs(
        ["Balance Sheet", "P&L", "Cash Flow", "SOCE", "Ageing", "Validation"]
    )

    with tab_bs:
        for section_name, section in [("Equity and Liabilities", bs.equity_and_liabilities), ("Assets", bs.assets)]:
            st.markdown(f"**{section_name}**")
            for major in section:
                st.write(f"*{major.major_head}*")
                for sh in major.sub_heads:
                    st.write(f"&nbsp;&nbsp;{sh.sub_head}: ₹{sh.current_year:,.2f} (PY ₹{sh.previous_year:,.2f})",
                              unsafe_allow_html=True)

    with tab_pnl:
        st.write(f"Total Revenue: ₹{pnl.total_revenue_cy:,.2f}")
        st.write(f"Total Expenses: ₹{pnl.total_expenses_cy:,.2f}")
        st.write(f"Profit After Tax: ₹{pnl.profit_after_tax_cy:,.2f}")

    with tab_cf:
        st.write(f"Net Cash from Operating Activities: ₹{cash_flow.cash_from_operations:,.2f}")
        st.write(f"Net Cash from Investing Activities: ₹{cash_flow.cash_from_investing:,.2f}")
        st.write(f"Net Cash from Financing Activities: ₹{cash_flow.cash_from_financing:,.2f}")

    with tab_soce:
        st.write(f"Equity Share Capital -- Closing: ₹{soce.equity_share_capital.closing:,.2f}")
        for comp in soce.other_equity:
            st.write(f"{comp.component}: Opening ₹{comp.opening:,.2f} → Closing ₹{comp.closing:,.2f}")

    with tab_ageing:
        for label, grid in [("Trade Receivables", receivables_ageing), ("Trade Payables", payables_ageing)]:
            st.markdown(f"**{label}**")
            if not grid.available:
                st.info(grid.unavailable_reason)
            else:
                st.write(f"Total: ₹{grid.total:,.2f} -- Reconciles: {'✅' if grid.reconciles_to_balance_sheet else '❌'}")

    with tab_val:
        if not issues:
            st.success("No validation issues.")
        for issue in issues:
            fn = {"ERROR": st.error, "WARNING": st.warning, "INFO": st.info}.get(issue.severity, st.write)
            fn(f"[{issue.code}] {issue.message}")

    # =======================================================================
    # INCOME TAX MODULE -- Income-tax Act, 2025 (Tax Year 2026-27)
    # =======================================================================
    st.divider()
    st.subheader("4. Income Tax — Computation, Deferred Tax & Tax Audit (Income-tax Act, 2025)")
    st.caption(
        "Auto-computed from the statements above; every figure below can be manually "
        "overridden before export. Rates per the Income-tax Act, 2025 / Finance Act, 2026 — "
        "verify against the enacted text before filing."
    )

    tc1, tc2 = st.columns(2)
    entity_type = tc1.selectbox("Entity type", list(EntityType), format_func=lambda e: e.value)
    regime = tc2.selectbox("Tax regime", REGIMES_BY_ENTITY[entity_type], format_func=lambda r: r.value)

    tax_tab_comp, tax_tab_dep, tax_tab_dt, tax_tab_audit = st.tabs(
        ["Income Computation", "IT Depreciation", "Deferred Tax", "Tax Audit (Form 26)"]
    )

    # ---- IT Depreciation (editable block schedule) ----
    with tax_tab_dep:
        st.markdown("**Depreciation as per the Income-tax Act (Block of Assets — WDV)**")
        st.caption(
            "Seeded from the Trial Balance (PY closing → opening WDV; CY increase → additions). "
            "Correct the ≥180/<180-day split, deletions, and opening WDV per the last filed return."
        )
        if "it_dep_rows" not in st.session_state:
            seeded = seed_schedule_from_tb(tb, mappings)
            st.session_state.it_dep_rows = [
                {"Block": r.block_name, "Rate": r.rate, "Opening WDV": r.opening_wdv,
                 "Additions ≥180 days": r.additions_180_plus,
                 "Additions <180 days": r.additions_less_180, "Deletions": r.deletions}
                for r in seeded.rows
            ] or [{"Block": "Plant and Machinery - General", "Rate": 0.15, "Opening WDV": 0.0,
                    "Additions ≥180 days": 0.0, "Additions <180 days": 0.0, "Deletions": 0.0}]
        edited = st.data_editor(
            st.session_state.it_dep_rows, num_rows="dynamic", use_container_width=True,
            column_config={
                "Block": st.column_config.SelectboxColumn(options=list(IT_DEPRECIATION_BLOCKS.keys())),
                "Rate": st.column_config.NumberColumn(min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
            }, key="it_dep_editor",
        )
        st.session_state.it_dep_rows = edited
        it_dep = ITDepreciationSchedule(rows=[
            AssetBlockRow(
                block_name=row.get("Block") or "Plant and Machinery - General",
                rate=float(row.get("Rate") or IT_DEPRECIATION_BLOCKS.get(row.get("Block", ""), 0.15)),
                opening_wdv=float(row.get("Opening WDV") or 0),
                additions_180_plus=float(row.get("Additions ≥180 days") or 0),
                additions_less_180=float(row.get("Additions <180 days") or 0),
                deletions=float(row.get("Deletions") or 0),
            ) for row in edited
        ])
        d1, d2 = st.columns(2)
        d1.metric("Total IT Depreciation", f"₹{it_dep.total_depreciation:,.2f}")
        d2.metric("Closing WDV (all blocks)", f"₹{it_dep.total_closing_wdv:,.2f}")

    # ---- Income Computation (auto + overrides) ----
    with tax_tab_comp:
        gratuity_unpaid = st.number_input("Unpaid gratuity provision (auto add-back)", value=0.0, step=1000.0)
        leave_unpaid = st.number_input("Unpaid leave encashment provision (auto add-back)", value=0.0, step=1000.0)
        comp = build_standard_computation(entity_type, regime, pnl, it_dep,
                                           gratuity_provision_unpaid=gratuity_unpaid,
                                           leave_encashment_unpaid=leave_unpaid)

        c1, c2 = st.columns(2)
        comp.brought_forward_business_loss = c1.number_input("B/f business loss (per last return)", value=0.0, step=1000.0)
        comp.unabsorbed_depreciation_bf = c2.number_input("B/f unabsorbed depreciation", value=0.0, step=1000.0)
        if entity_type == EntityType.COMPANY and regime == TaxRegime.COMPANY_NORMAL:
            comp.turnover_reference_year = st.number_input(
                "Turnover in the reference year (for 25% vs 30% rate)", value=float(pnl.total_revenue_cy), step=100000.0)

        st.markdown("**Adjustments** — auto figures shown; tick *Override* to replace any line manually.")
        for adj in comp.additions + comp.deductions:
            oc1, oc2, oc3 = st.columns([4, 1, 2])
            oc1.write(f"{adj.label}  \n:gray[{adj.section_ref}]")
            use_override = oc2.checkbox("Override", key=f"ov_{adj.code}")
            if use_override:
                adj.override = oc3.number_input("Amount", value=float(adj.auto), step=100.0,
                                                key=f"ovamt_{adj.code}", label_visibility="collapsed")
            else:
                oc3.write(f"₹{adj.auto:,.2f}")

        p1, p2, p3, p4 = st.columns(4)
        comp.tds_credit = p1.number_input("TDS credit", value=0.0, step=1000.0)
        comp.tcs_credit = p2.number_input("TCS credit", value=0.0, step=1000.0)
        comp.advance_tax_paid = p3.number_input("Advance tax paid", value=0.0, step=1000.0)
        comp.self_assessment_tax_paid = p4.number_input("Self-assessment tax", value=0.0, step=1000.0)

        months = st.slider("Months from year-end to expected filing (for interest estimate)", 0, 12, 6)
        comp.auto_advance_tax_interest(months)
        for iline in (comp.interest_234a, comp.interest_234b, comp.interest_234c):
            ic1, ic2, ic3 = st.columns([4, 1, 2])
            ic1.write(iline.label)
            if ic2.checkbox("Override", key=f"ov_{iline.code}"):
                iline.override = ic3.number_input("Amount", value=float(iline.auto), step=100.0,
                                                  key=f"ovamt_{iline.code}", label_visibility="collapsed")
            else:
                ic3.write(f"₹{iline.auto:,.2f}")

        st.markdown("---")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Taxable Income", f"₹{comp.taxable_income:,.0f}")
        r2.metric("Total Tax Liability", f"₹{comp.total_tax_liability:,.0f}",
                  help="Tax + surcharge (with marginal relief) + 4% cess")
        r3.metric("Prepaid Taxes", f"₹{comp.prepaid_taxes:,.0f}")
        r4.metric("Net Payable / (Refund)", f"₹{comp.net_payable:,.0f}")
        st.session_state.tax_comp = comp

    # ---- Deferred Tax ----
    with tax_tab_dt:
        st.markdown("**Deferred Tax Working (AS 22 pattern)**")
        book_wdv_auto = bs_amount("Property, Plant and Equipment") + bs_amount("Intangible Assets")
        dt1, dt2 = st.columns(2)
        book_wdv = dt1.number_input("Book WDV of fixed assets (net block)", value=float(book_wdv_auto), step=1000.0)
        if entity_type == EntityType.COMPANY:
            base_for_dt = {TaxRegime.COMPANY_CONCESSIONAL_22: 0.22,
                           TaxRegime.COMPANY_NEW_MFG_15: 0.15}.get(regime, 0.25)
            sur_for_dt = COMPANY_CONCESSIONAL_SURCHARGE if regime in (
                TaxRegime.COMPANY_CONCESSIONAL_22, TaxRegime.COMPANY_NEW_MFG_15) else 0.07
        elif entity_type == EntityType.FIRM_LLP:
            base_for_dt, sur_for_dt = 0.30, 0.0
        else:
            base_for_dt, sur_for_dt = 0.30, 0.0
        eff_rate = dt2.number_input("Effective tax rate for DTA/DTL", value=effective_tax_rate(base_for_dt, sur_for_dt),
                                    step=0.001, format="%.5f",
                                    help="e.g. 22% × 1.10 surcharge × 1.04 cess = 0.25168")
        g1, g2, g3 = st.columns(3)
        dt_gratuity = g1.number_input("Gratuity provision (books)", value=gratuity_unpaid, step=1000.0)
        dt_leave = g2.number_input("Leave encashment provision (books)", value=leave_unpaid, step=1000.0)
        dt_tds = g3.number_input("TDS-default disallowance c/f", value=0.0, step=1000.0)
        dt_calc = build_standard_deferred_tax(eff_rate, it_dep.total_closing_wdv, book_wdv,
                                              gratuity_provision=dt_gratuity,
                                              leave_encashment_provision=dt_leave,
                                              tds_default_disallowance=dt_tds)
        rows_view = [{
            "Particulars": i.particulars, "As per IT": i.as_per_income_tax, "As per Books": i.as_per_books,
            "Timing Diff.": i.timing_difference, "Closing DTA/(DTL)": dt_calc.closing_for(i),
        } for i in dt_calc.items]
        st.dataframe(rows_view, use_container_width=True)
        m1, m2 = st.columns(2)
        m1.metric("Closing DTA/(DTL)", f"₹{dt_calc.total_closing:,.2f}")
        m2.metric("Charge/(Credit) for the year", f"₹{dt_calc.total_movement:,.2f}",
                  help="Flows to P&L as deferred tax; positive = DTA created (credit)")
        st.session_state.dt_calc = dt_calc
        st.session_state.book_wdv = book_wdv

    # ---- Tax Audit ----
    with tax_tab_audit:
        st.markdown("**Applicability — s.63, Income-tax Act, 2025**")
        a1, a2, a3 = st.columns(3)
        is_prof = a1.checkbox("Profession (not business)?", value=False)
        turnover_in = a2.number_input("Turnover / gross receipts",
                                      value=float(pnl.total_revenue_cy), step=100000.0)
        a3.write("")
        c1, c2 = st.columns(2)
        cash_rec = c1.slider("Cash receipts as % of total", 0, 100, 3) / 100
        cash_pay = c2.slider("Cash payments as % of total", 0, 100, 3) / 100
        pr1, pr2 = st.columns(2)
        presumptive = pr1.checkbox("Opted a presumptive scheme?")
        below_deemed = pr2.checkbox("Declaring below deemed profit rate?") if presumptive else False

        applicability = check_applicability(entity_type, is_prof, turnover_in,
                                            cash_rec, cash_pay, presumptive, below_deemed)
        if applicability.required:
            st.error("TAX AUDIT REQUIRED — report in " + applicability.form)
        else:
            st.success("Tax audit not required on these facts.")
        for reason in applicability.reasons:
            st.write("• " + reason)
        if applicability.required:
            st.caption(applicability.due_note +
                       f" Fee for default u/s 446: ₹{applicability.fee_if_defaulted:,.0f}.")

        st.markdown("**Form 26, Part B — clause checklist** "
                    ":gray[(anchors 49–51 confirmed; registry updatable as CBDT finalizes)]")
        checklist = build_form26_checklist({
            "company_name": company.name, "pan": company.pan, "gstin": company.gstin,
            "registered_office": company.registered_office,
            "entity_status": entity_type.value, "tax_year": "2026-27",
            "turnover": f"CY ₹{pnl.total_revenue_cy:,.0f} / PY ₹{pnl.total_revenue_py:,.0f}",
            "it_depreciation": f"{len(it_dep.rows)} blocks; dep ₹{it_dep.total_depreciation:,.0f}; closing WDV ₹{it_dep.total_closing_wdv:,.0f}",
            "gratuity_disallowance": (f"₹{gratuity_unpaid:,.0f} added back (unpaid provision)"
                                       if gratuity_unpaid else ""),
            "cfl_schedule": (f"B/f loss ₹{comp.brought_forward_business_loss:,.0f}; "
                             f"unabs. dep ₹{comp.unabsorbed_depreciation_bf:,.0f}"),
            "ratios": "; ".join(f"{ra.name}: {ra.current_year:.2f}" for ra in ratios[:4]
                                 if ra.current_year is not None),
            "auditor": company.auditor,
        })
        done = sum(1 for c in checklist if c.status == "Completed")
        st.progress(done / len(checklist), text=f"{done}/{len(checklist)} clauses pre-filled by engine")
        for item in checklist:
            with st.expander(f"Clause {item.no}: {item.title}"
                             + (" ✅" if item.status == "Completed" else "")):
                if item.old_3cd_ref:
                    st.caption(f"Old-form reference: {item.old_3cd_ref}")
                if item.guidance:
                    st.info(item.guidance)
                if item.auto_value:
                    st.write(f"**Engine pre-fill:** {item.auto_value}")
                item.response = st.text_area("Auditor response / particulars",
                                             key=f"cl_{item.no}", height=68)
                if item.response:
                    item.status = "Completed"
        st.session_state.audit_applicability = applicability
        st.session_state.audit_checklist = checklist

    # ---- Tax downloads (workbook including tax sheets + computation PDF) ----
    st.markdown("### Downloads — with Income Tax module")
    wb_tax = build_workbook(company, fy_label, tb, mappings, bs, pnl, cash_flow, notes, issues, ratios,
                             soce=soce, receivables_ageing=receivables_ageing, payables_ageing=payables_ageing)
    add_tax_sheets(wb_tax, company, comp, it_dep, dt_calc, st.session_state.book_wdv,
                   applicability, checklist)
    tax_excel_buf = io.BytesIO()
    wb_tax.save(tax_excel_buf)
    tax_excel_buf.seek(0)

    comp_pdf_tmp = "/tmp/_Computation_of_Income.pdf"
    build_computation_pdf(comp_pdf_tmp, company, comp, it_dep, dt_calc, applicability)
    with open(comp_pdf_tmp, "rb") as f:
        comp_pdf_bytes = f.read()

    td1, td2 = st.columns(2)
    td1.download_button("⬇ Excel — Financials + Tax Comp + Deferred Tax + IT Dep + Audit",
                         data=tax_excel_buf,
                         file_name=f"{company.name}_Financials_and_Tax.xlsx",
                         use_container_width=True)
    td2.download_button("⬇ PDF — Computation of Total Income (paper return working)",
                         data=comp_pdf_bytes,
                         file_name=f"{company.name}_Computation_of_Income.pdf",
                         use_container_width=True)

    st.divider()
    if st.button("⬅ Start a new Trial Balance"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
