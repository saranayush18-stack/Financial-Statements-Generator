"""
PDF Export.

Produces a professional, print-ready financial statement PDF: cover page,
Balance Sheet, P&L, Cash Flow, Notes to Accounts, Accounting Policies, with
page numbers, header/footer, and a signature block -- structured to look
like an audited financial statement package, not a raw data dump.
"""
from __future__ import annotations

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, NextPageTemplate
)

from models import Company
from core.statement_generator import BalanceSheet, ProfitAndLoss, CashFlowStatement
from core.notes_generator import Note, STANDARD_ACCOUNTING_POLICIES
from core.soce_generator import StatementOfChangesInEquity
from core.ageing import AgeingGrid, BUCKET_LABELS, CATEGORY_LABELS

NAVY = colors.HexColor("#1F4E79")
LIGHT_BLUE = colors.HexColor("#D9E1F2")
GREY = colors.HexColor("#808080")


def _styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("CompanyTitle", parent=styles["Title"], fontName="Times-Bold", fontSize=20,
                               textColor=NAVY, alignment=TA_CENTER, spaceAfter=6))
    styles.add(ParagraphStyle("StatementTitle", parent=styles["Heading1"], fontName="Times-Bold", fontSize=14,
                               alignment=TA_CENTER, textColor=colors.black, spaceAfter=4))
    styles.add(ParagraphStyle("SubTitle", parent=styles["Normal"], fontName="Times-Italic", fontSize=10,
                               alignment=TA_CENTER, textColor=GREY, spaceAfter=2))
    styles.add(ParagraphStyle("SectionHeader", parent=styles["Heading2"], fontName="Times-Bold", fontSize=11,
                               textColor=NAVY, spaceBefore=10, spaceAfter=4))
    styles.add(ParagraphStyle("NoteHeader", parent=styles["Heading3"], fontName="Times-Bold", fontSize=10,
                               textColor=colors.white, backColor=NAVY, spaceBefore=8, spaceAfter=2, leftIndent=2))
    styles.add(ParagraphStyle("Body", parent=styles["Normal"], fontName="Times-Roman", fontSize=9, leading=13))
    styles.add(ParagraphStyle("PolicyTitle", parent=styles["Heading3"], fontName="Times-Bold", fontSize=10,
                               textColor=NAVY))
    return styles


def _fmt(n: float) -> str:
    if n is None:
        return "-"
    if abs(n) < 0.005:
        return "-"
    neg = n < 0
    s = f"{abs(n):,.2f}"
    return f"({s})" if neg else s


def _header_footer(canvas, doc, company: Company, fy_label: str):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(GREY)
    canvas.drawString(20 * mm, 12 * mm, f"{company.name} — Financial Statements for {fy_label}")
    canvas.drawRightString(190 * mm, 12 * mm, f"Page {doc.page}")
    canvas.setStrokeColor(LIGHT_BLUE)
    canvas.line(20 * mm, 15 * mm, 190 * mm, 15 * mm)
    canvas.restoreState()


SERIF = "Times-Roman"
SERIF_BOLD = "Times-Bold"


def _note_no(note_ref: str | None) -> str:
    import re
    if not note_ref:
        return ""
    m = re.match(r"Note\s+(\d+)", note_ref)
    return m.group(1) if m else ""


def _period_labels(company: Company) -> tuple[str, str]:
    cy_end = company.financial_year_end
    try:
        py_end = cy_end.replace(year=cy_end.year - 1)
    except ValueError:
        py_end = cy_end.replace(year=cy_end.year - 1, day=28)
    fmt = "%d %B %Y"
    return cy_end.strftime(fmt), py_end.strftime(fmt)


def _schedule3_table(rows, col_widths, bold_rows=None, box_rows=None):
    """Statutory-format table: Times New Roman, boxed header, bold subtotals,
    double rule under grand-total rows -- matching the Excel Schedule III
    output so PDF and Excel look like the same signed financial statement."""
    bold_rows = bold_rows or set()
    box_rows = box_rows or set()
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), SERIF),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), SERIF_BOLD),
        ("BOX", (0, 0), (-1, 0), 1, colors.black),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
    ]
    for i in bold_rows:
        style_cmds.append(("FONTNAME", (0, i), (-1, i), SERIF_BOLD))
    for i in box_rows:
        style_cmds.append(("LINEABOVE", (0, i), (-1, i), 0.5, colors.black))
        style_cmds.append(("LINEBELOW", (0, i), (-1, i), 1.75, colors.black))
    t = Table(rows, colWidths=col_widths)
    t.setStyle(TableStyle(style_cmds))
    return t


def _signature_block_pdf(company: Company, styles) -> list:
    d1 = company.directors[0] if len(company.directors) > 0 else "Director"
    d2 = company.directors[1] if len(company.directors) > 1 else "Director"
    left = (
        "For and on behalf of the Board of Directors<br/><br/><br/><br/>"
        f"{d1}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{d2}<br/>"
        "<i>Director&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Director</i><br/>"
        "DIN:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;DIN:<br/><br/>"
        "Place:<br/>Date:<br/>UDIN:"
    )
    right = (
        "As per our report of even date attached<br/>"
        f"For {company.auditor or '&lt;Auditor Firm Name&gt;'}<br/>"
        "Chartered Accountants<br/>Firm Registration No.:<br/><br/><br/><br/>"
        "Partner<br/>Membership No.:"
    )
    tbl = Table([[Paragraph(left, styles["Body"]), Paragraph(right, styles["Body"])]],
                colWidths=[95 * mm, 75 * mm])
    tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return [Spacer(1, 4),
            Paragraph("See accompanying notes forming part of the financial statements",
                      ParagraphStyle("Ital", parent=styles["Body"], fontName="Times-Italic")),
            Spacer(1, 10), tbl]


def _section_table(rows, col_widths, header_row: bool = True):
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
    ]
    if header_row:
        style_cmds += [
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ]
    t = Table(rows, colWidths=col_widths)
    t.setStyle(TableStyle(style_cmds))
    return t


def _build_balance_sheet_flowables(bs: BalanceSheet, company: Company, styles) -> list:
    cy_label, py_label = _period_labels(company)
    flow = [Paragraph("Balance Sheet", styles["StatementTitle"]),
            Paragraph(f"as at {cy_label}", styles["SubTitle"]), Spacer(1, 8)]
    header = ["Sl.\nNo.", "Particulars", "Note\nNo.", f"As at\n{cy_label}", f"As at\n{py_label}"]
    rows = [header]
    bold_idx: set[int] = set()
    box_idx: set[int] = set()

    def add_group(major_list, numeral, title):
        rows.append([numeral, title, "", "", ""])
        bold_idx.add(len(rows) - 1)
        for m_idx, major in enumerate(major_list, start=1):
            rows.append(["", f"({m_idx}) {major.major_head}", "", "", ""])
            bold_idx.add(len(rows) - 1)
            for s_idx, sh in enumerate(major.sub_heads):
                letter = chr(ord('a') + s_idx)
                rows.append(["", f"     ({letter}) {sh.sub_head}", _note_no(sh.note_ref),
                              _fmt(sh.current_year), _fmt(sh.previous_year)])
            rows.append(["", f"Total - {major.major_head}", "",
                          _fmt(major.current_year), _fmt(major.previous_year)])
            bold_idx.add(len(rows) - 1)

    add_group(bs.equity_and_liabilities, "I", "EQUITY AND LIABILITIES")
    rows.append(["", "TOTAL - EQUITY AND LIABILITIES", "",
                  _fmt(bs.total_equity_and_liabilities_cy), _fmt(bs.total_equity_and_liabilities_py)])
    bold_idx.add(len(rows) - 1)
    box_idx.add(len(rows) - 1)

    add_group(bs.assets, "II", "ASSETS")
    rows.append(["", "TOTAL - ASSETS", "", _fmt(bs.total_assets_cy), _fmt(bs.total_assets_py)])
    bold_idx.add(len(rows) - 1)
    box_idx.add(len(rows) - 1)

    flow.append(_schedule3_table(rows, [12 * mm, 78 * mm, 15 * mm, 32 * mm, 32 * mm], bold_idx, box_idx))
    flow.append(Spacer(1, 6))
    tie_text = "Balance Sheet tallies." if bs.is_tallied else "WARNING: Balance Sheet does not tally — review mapping."
    flow.append(Paragraph(tie_text, styles["Body"]))
    flow += _signature_block_pdf(company, styles)
    return flow


def _build_pnl_flowables(pnl: ProfitAndLoss, company: Company, styles) -> list:
    cy_label, py_label = _period_labels(company)
    flow = [Paragraph("Statement of Profit and Loss", styles["StatementTitle"]),
            Paragraph(f"for the year ended {cy_label}", styles["SubTitle"]), Spacer(1, 8)]
    header = ["Sl.\nNo.", "Particulars", "Note\nNo.", f"Year ended\n{cy_label}", f"Year ended\n{py_label}"]
    rows = [header]
    bold_idx: set[int] = set()
    box_idx: set[int] = set()

    rows.append(["I", "Revenue", "", "", ""])
    bold_idx.add(len(rows) - 1)
    for r in pnl.revenue:
        rows.append(["", f"     {r.sub_head}", _note_no(r.note_ref), _fmt(r.current_year), _fmt(r.previous_year)])
    rows.append(["", "Total Revenue (I)", "", _fmt(pnl.total_revenue_cy), _fmt(pnl.total_revenue_py)])
    bold_idx.add(len(rows) - 1)

    rows.append(["II", "Expenses", "", "", ""])
    bold_idx.add(len(rows) - 1)
    for e in pnl.expenses:
        rows.append(["", f"     {e.sub_head}", _note_no(e.note_ref), _fmt(e.current_year), _fmt(e.previous_year)])
    rows.append(["", "Total Expenses (II)", "", _fmt(pnl.total_expenses_cy), _fmt(pnl.total_expenses_py)])
    bold_idx.add(len(rows) - 1)

    rows.append(["III", "Profit Before Tax (I - II)", "",
                  _fmt(pnl.profit_before_tax_cy), _fmt(pnl.profit_before_tax_py)])
    bold_idx.add(len(rows) - 1)
    rows.append(["IV", "Tax Expense", "", _fmt(pnl.tax_expense_cy), _fmt(pnl.tax_expense_py)])
    rows.append(["V", "Profit After Tax (III - IV)", "",
                  _fmt(pnl.profit_after_tax_cy), _fmt(pnl.profit_after_tax_py)])
    bold_idx.add(len(rows) - 1)
    box_idx.add(len(rows) - 1)

    flow.append(_schedule3_table(rows, [12 * mm, 78 * mm, 15 * mm, 32 * mm, 32 * mm], bold_idx, box_idx))
    flow += _signature_block_pdf(company, styles)
    return flow


def _build_soce_flowables(soce: StatementOfChangesInEquity, styles) -> list:
    flow = [Paragraph("Statement of Changes in Equity", styles["StatementTitle"]), Spacer(1, 6)]

    flow.append(Paragraph("A. Equity Share Capital", styles["SectionHeader"]))
    sc = soce.equity_share_capital
    sc_rows = [
        ["Particulars", "Amount"],
        ["Balance at the beginning of the year", _fmt(sc.opening)],
        ["Changes in equity share capital during the year", _fmt(sc.changes_during_year)],
        ["Balance at the end of the year", _fmt(sc.closing)],
    ]
    flow.append(_section_table(sc_rows, [130 * mm, 40 * mm]))
    flow.append(Spacer(1, 10))

    flow.append(Paragraph("B. Other Equity", styles["SectionHeader"]))
    oe_rows = [["Component", "Opening", "Profit for Year", "Other Movements", "Closing"]]
    for comp in soce.other_equity:
        oe_rows.append([
            comp.component, _fmt(comp.opening), _fmt(comp.profit_for_the_year),
            _fmt(comp.other_movements), _fmt(comp.closing),
        ])
    oe_rows.append([
        "Total Other Equity", _fmt(soce.total_other_equity_opening), "", "",
        _fmt(soce.total_other_equity_closing),
    ])
    flow.append(_section_table(oe_rows, [55 * mm, 32 * mm, 32 * mm, 32 * mm, 32 * mm]))
    flow.append(Spacer(1, 6))
    flow.append(Paragraph(
        "Note: \"Other Movements\" is a residual figure (closing less opening less profit "
        "transferred) capturing dividends paid, inter-reserve transfers, or prior period "
        "adjustments -- a Trial Balance alone cannot distinguish these; verify against board "
        "resolutions and minutes before finalizing.", styles["Body"]))
    return flow


def _build_ageing_flowables(receivables: AgeingGrid | None, payables: AgeingGrid | None, styles) -> list:
    flow = [Paragraph("Ageing Schedule - Trade Receivables and Trade Payables", styles["StatementTitle"]),
            Spacer(1, 6)]

    def render_grid(grid: AgeingGrid, title: str):
        flow.append(Paragraph(title, styles["SectionHeader"]))
        if not grid.available:
            flow.append(Paragraph(
                f"Ageing schedule not available: {grid.unavailable_reason}", styles["Body"]))
            flow.append(Spacer(1, 8))
            return
        header = ["Category"] + [b.replace(" - ", "-\n") for b in BUCKET_LABELS] + ["Total"]
        rows = [header]
        for cat in CATEGORY_LABELS:
            row_total = sum(grid.grid[cat][b] for b in BUCKET_LABELS)
            rows.append([cat] + [_fmt(grid.grid[cat][b]) for b in BUCKET_LABELS] + [_fmt(row_total)])
        rows.append(["Total"] + ["" for _ in BUCKET_LABELS] + [_fmt(grid.total)])
        col_widths = [38 * mm] + [17 * mm] * len(BUCKET_LABELS) + [18 * mm]
        flow.append(_section_table(rows, col_widths))
        if grid.reconciles_to_balance_sheet is False:
            flow.append(Paragraph(
                f"WARNING: Ageing total ({grid.total:,.2f}) does not reconcile to the Balance "
                f"Sheet figure ({grid.balance_sheet_amount:,.2f}).",
                ParagraphStyle("Warn", parent=styles["Body"], textColor=colors.red)))
        flow.append(Spacer(1, 10))

    if receivables is not None:
        render_grid(receivables, "Trade Receivables Ageing")
    if payables is not None:
        render_grid(payables, "Trade Payables Ageing")
    return flow


def _build_cash_flow_flowables(cf: CashFlowStatement, styles) -> list:
    flow = [Paragraph("Cash Flow Statement (Indirect Method)", styles["StatementTitle"])]
    rows = [["Particulars", "Amount"]]
    rows.append(["A. Cash Flow from Operating Activities", ""])
    rows.append(["Net Profit Before Tax", _fmt(cf.net_profit_before_tax)])
    rows.append(["Add: Depreciation", _fmt(cf.depreciation_addback)])
    rows.append(["Add: Interest Expense", _fmt(cf.interest_expense_addback)])
    for label, val in cf.working_capital_changes.items():
        rows.append([f"(Increase)/Decrease in {label}", _fmt(val)])
    rows.append(["Net Cash from Operating Activities", _fmt(cf.cash_from_operations)])
    rows.append(["B. Net Cash used in Investing Activities", _fmt(cf.cash_from_investing)])
    rows.append(["C. Net Cash from/used in Financing Activities", _fmt(cf.cash_from_financing)])
    rows.append(["Net Increase/(Decrease) in Cash (A+B+C)", _fmt(cf.net_increase_in_cash)])
    rows.append(["Cash at Beginning of the Year", _fmt(cf.opening_cash)])
    rows.append(["Cash at End of the Year", _fmt(cf.closing_cash)])
    flow.append(_section_table(rows, [130 * mm, 40 * mm]))
    return flow


def _build_notes_flowables(notes: list[Note], styles) -> list:
    flow = [Paragraph("Notes forming part of the Financial Statements", styles["StatementTitle"]), Spacer(1, 6)]
    for note in notes:
        flow.append(Paragraph(note.note_ref, styles["NoteHeader"]))
        rows = [["Particulars", "Current Year", "Previous Year"]]
        for item in note.line_items:
            rows.append([item.label, _fmt(item.current_year), _fmt(item.previous_year)])
        total_row = len(rows)
        rows.append(["Total", _fmt(note.total_current_year), _fmt(note.total_previous_year)])
        flow.append(_schedule3_notelike_table(rows, [90 * mm, 40 * mm, 40 * mm], {total_row}))
        flow.append(Spacer(1, 8))
    return flow


def _schedule3_notelike_table(rows, col_widths, bold_rows):
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), SERIF),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), SERIF_BOLD),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
    ]
    for i in bold_rows:
        style_cmds.append(("FONTNAME", (0, i), (-1, i), SERIF_BOLD))
        style_cmds.append(("LINEABOVE", (0, i), (-1, i), 1, colors.black))
    t = Table(rows, colWidths=col_widths)
    t.setStyle(TableStyle(style_cmds))
    return t


def _build_policies_flowables(styles) -> list:
    flow = [Paragraph("Significant Accounting Policies", styles["StatementTitle"]), Spacer(1, 6)]
    for i, (title, text) in enumerate(STANDARD_ACCOUNTING_POLICIES.items(), start=1):
        flow.append(Paragraph(f"{i}. {title}", styles["PolicyTitle"]))
        flow.append(Paragraph(text, styles["Body"]))
        flow.append(Spacer(1, 6))
    return flow


def _build_cover_page(company: Company, fy_label: str, styles) -> list:
    flow = [
        Spacer(1, 60 * mm),
        Paragraph(company.name, styles["CompanyTitle"]),
        Spacer(1, 10),
        Paragraph("Financial Statements", styles["StatementTitle"]),
        Paragraph(f"For the Year Ended {fy_label}", styles["SubTitle"]),
        Spacer(1, 20),
    ]
    meta_rows = []
    if company.cin:
        meta_rows.append(["CIN", company.cin])
    if company.pan:
        meta_rows.append(["PAN", company.pan])
    if company.gstin:
        meta_rows.append(["GSTIN", company.gstin])
    if company.registered_office:
        meta_rows.append(["Registered Office", company.registered_office])
    if meta_rows:
        t = Table(meta_rows, colWidths=[40 * mm, 110 * mm])
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ]))
        flow.append(t)
    flow.append(Spacer(1, 40 * mm))
    if company.auditor or company.directors:
        sig_rows = [["For and on behalf of the Board of Directors", ""]]
        if company.directors:
            for d in company.directors:
                sig_rows.append([d, "Director"])
        if company.auditor:
            sig_rows.append(["", ""])
            sig_rows.append([company.auditor, "Auditor"])
        t2 = Table(sig_rows, colWidths=[90 * mm, 60 * mm])
        t2.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("TOPPADDING", (0, 0), (-1, -1), 14),
        ]))
        flow.append(t2)
    return flow


def build_pdf(
    output_path: str,
    company: Company,
    fy_label: str,
    bs: BalanceSheet,
    pnl: ProfitAndLoss,
    cash_flow: CashFlowStatement,
    notes: list[Note],
    soce: StatementOfChangesInEquity | None = None,
    receivables_ageing: AgeingGrid | None = None,
    payables_ageing: AgeingGrid | None = None,
) -> str:
    styles = _styles()
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        topMargin=20 * mm, bottomMargin=20 * mm, leftMargin=20 * mm, rightMargin=20 * mm,
        title=f"{company.name} - Financial Statements {fy_label}",
    )

    story = []
    story += _build_cover_page(company, fy_label, styles)
    story.append(PageBreak())
    story += _build_balance_sheet_flowables(bs, company, styles)
    story.append(PageBreak())
    story += _build_pnl_flowables(pnl, company, styles)
    if soce is not None:
        story.append(PageBreak())
        story += _build_soce_flowables(soce, styles)
    story.append(PageBreak())
    story += _build_cash_flow_flowables(cash_flow, styles)
    if receivables_ageing is not None or payables_ageing is not None:
        story.append(PageBreak())
        story += _build_ageing_flowables(receivables_ageing, payables_ageing, styles)
    story.append(PageBreak())
    story += _build_notes_flowables(notes, styles)
    story.append(PageBreak())
    story += _build_policies_flowables(styles)

    def on_page(canvas, doc_):
        _header_footer(canvas, doc_, company, fy_label)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return output_path
