#!/usr/bin/env python3
"""
Build Application-Risk-Assessment-Template.xlsx

An open-source Likelihood x Impact (5x5) application risk assessment workbook,
aligned with ISO/IEC 27001:2022 (clause 6.1.2), NIST CSF 2.0 and ISO/IEC 42001.

Usage:
    pip install openpyxl
    python scripts/build_template.py            # writes ../Application-Risk-Assessment-Template.xlsx

The workbook contains no macros and no personal data. All sample entries are
fictional. Licensed under the MIT License (see LICENSE).
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------- #
# Design tokens
# --------------------------------------------------------------------------- #
FONT = "Arial"
NAVY = "1F3864"          # dark blue headers
MID_BLUE = "2F5597"
LIGHT_BLUE = "D9E2F3"
PALE = "F2F2F2"
INPUT_BLUE = "0000FF"    # classic modelling convention: blue = input
GRID = "BFBFBF"

# (label, min score, max score, fill, text colour)
BANDS = [
    ("Low", 1, 4, "C6E0B4", "1E4620"),
    ("Medium", 5, 9, "FFE699", "7F6000"),
    ("High", 10, 16, "F4B183", "7F3200"),
    ("Critical", 17, 25, "C00000", "FFFFFF"),
]

# Sheet names
S_METH = "01_Methodology"
S_SCALES = "02_Scales"
S_INV = "03_Application_Inventory"
S_REG = "04_Risk_Register"
S_DASH = "05_Dashboard"
S_MATRIX = "06_Risk_Matrix"

# Data ranges
INV_FIRST, INV_LAST = 5, 34      # inventory rows available (30 applications)
REG_FIRST, REG_LAST = 5, 50      # risk register rows available (46 risks)

thin = Side(style="thin", color=GRID)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


# --------------------------------------------------------------------------- #
# Small style helpers
# --------------------------------------------------------------------------- #
def f(bold=False, color="000000", size=10, italic=False) -> Font:
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)


def put(ws, ref, value=None, *, font=None, bg=None, align=None, border=False, fmt=None):
    cell = ws[ref]
    if value is not None:
        cell.value = value
    cell.font = font or f()
    if bg:
        cell.fill = fill(bg)
    cell.alignment = align or Alignment(vertical="top", wrap_text=True)
    if border:
        cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    return cell


def header(ws, row, labels, *, start_col=1, height=34, bg=NAVY):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=label)
        c.font = f(bold=True, color="FFFFFF")
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height


def title(ws, text, subtitle=None, width_cols=8):
    ws["A1"].value = text
    ws["A1"].font = f(bold=True, color="FFFFFF", size=16)
    ws["A1"].alignment = Alignment(vertical="center")
    for col in range(1, width_cols + 1):
        ws.cell(row=1, column=col).fill = fill(NAVY)
    ws.row_dimensions[1].height = 34
    if subtitle:
        ws["A2"].value = subtitle
        ws["A2"].font = f(italic=True, color="404040")
        ws["A2"].alignment = Alignment(vertical="center")
        ws.row_dimensions[2].height = 22


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def page_setup(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


def band_for(score: int):
    for name, lo, hi, bg, fg in BANDS:
        if lo <= score <= hi:
            return name, bg, fg
    raise ValueError(score)


# --------------------------------------------------------------------------- #
# 01_Methodology
# --------------------------------------------------------------------------- #
def build_methodology(wb: Workbook):
    ws = wb.active
    ws.title = S_METH
    ws.sheet_properties.tabColor = NAVY
    title(
        ws,
        "Application Risk Assessment Template",
        "Likelihood × Impact (5×5)  |  ISO/IEC 27001:2022 cl. 6.1.2  |  NIST CSF 2.0  |  ISO/IEC 42001:2023",
        width_cols=2,
    )
    widths(ws, {"A": 30, "B": 118})
    ws.sheet_view.showGridLines = False

    ws["A1"].comment = Comment(
        "OPEN SOURCE: this template is released under the MIT License. You may freely use, copy, "
        "modify and redistribute it. It is provided 'as is' without warranty. All sample data is fictional.",
        "Template",
        width=360,
        height=110,
    )

    row = 3

    def section(text):
        nonlocal row
        row += 1
        c = ws.cell(row=row, column=1, value=text)
        c.font = f(bold=True, color="FFFFFF", size=11)
        for col in (1, 2):
            ws.cell(row=row, column=col).fill = fill(MID_BLUE)
        ws.row_dimensions[row].height = 22
        ws.cell(row=row, column=1).alignment = Alignment(vertical="center")
        row += 1

    def line(label, text, *, height=None, bold_label=True, bg=None, label_font=None):
        nonlocal row
        a = ws.cell(row=row, column=1, value=label)
        a.font = label_font or f(bold=bold_label)
        a.alignment = Alignment(vertical="top", wrap_text=True)
        b = ws.cell(row=row, column=2, value=text)
        b.font = f()
        b.alignment = Alignment(vertical="top", wrap_text=True)
        if bg:
            a.fill = fill(bg)
        if height:
            ws.row_dimensions[row].height = height
        row += 1

    ws["A3"].value = "Version 1.0  ·  Open source (MIT) — free to use, copy, modify and redistribute  ·  All sample data is fictional"
    ws["A3"].font = f(italic=True, color="595959")

    section("1. Purpose")
    line(
        "What it is",
        "A ready-to-use workbook to identify, score, treat and monitor information-security risks affecting applications — "
        "especially business-critical ones such as HRIS, ERP, CRM, customer portals and AI/ML engines.",
        height=34,
    )
    line(
        "Who it is for",
        "Security, GRC, IT and application owners who need a consistent, auditable risk register without buying a GRC platform.",
        height=20,
    )

    section("2. The method: Likelihood × Impact")
    line(
        "Formula",
        "Risk Score = Likelihood (1–5) × Impact (1–5)   →   a score from 1 to 25",
        label_font=f(bold=True),
        height=20,
    )
    line(
        "Reading the score",
        "Each risk is scored twice: INHERENT (before any treatment, considering only controls that are already unavoidable) and "
        "RESIDUAL (after existing and planned controls). The gap between the two shows the value of your controls.",
        height=34,
    )
    for name, lo, hi, bg, fg in BANDS:
        meaning = {
            "Low": "Acceptable within normal management. Monitor and review annually.",
            "Medium": "Treat or consciously accept with documented justification. Review twice a year.",
            "High": "Treatment plan with owner and date required. Acceptance needs senior management sign-off.",
            "Critical": "Immediate escalation to executive level. Interim controls now; full treatment plan within 30 days.",
        }[name]
        a = ws.cell(row=row, column=1, value=f"{name}  ({lo}–{hi})")
        a.font = f(bold=True, color=fg)
        a.fill = fill(bg)
        a.alignment = Alignment(horizontal="center", vertical="center")
        a.border = BORDER
        b = ws.cell(row=row, column=2, value=meaning)
        b.font = f()
        b.alignment = Alignment(vertical="center", wrap_text=True)
        b.border = BORDER
        ws.row_dimensions[row].height = 22
        row += 1

    section("3. Step-by-step instructions")
    steps = [
        ("Step 1 — Calibrate", "Read 02_Scales. Adjust the example money, downtime and data thresholds in the Impact table to fit your organisation's size and risk appetite."),
        ("Step 2 — Inventory", "List your applications in 03_Application_Inventory (up to 30 rows). Use the drop-downs. Note which are internet-facing and which contain AI/ML."),
        ("Step 3 — Identify risks", "In 04_Risk_Register, add one row per risk scenario. Choose the Application ID from the drop-down — the application name fills in automatically."),
        ("Step 4 — Describe clearly", "Write the scenario as: threat actor/event + weakness + consequence. Fill Threat, Vulnerability and Asset/Data affected separately."),
        ("Step 5 — Score inherent risk", "Pick Inherent Likelihood and Impact (1–5) using 02_Scales. Score and Rating are calculated automatically."),
        ("Step 6 — Decide treatment", "Choose Mitigate, Accept, Transfer or Avoid. Record existing/planned controls and map them to ISO 27001 Annex A and/or NIST CSF 2.0 (and ISO 42001 Annex A for AI)."),
        ("Step 7 — Score residual risk", "Re-score Likelihood and Impact assuming the controls are effective. Residual should not exceed inherent (it turns red if it does)."),
        ("Step 8 — Assign and track", "Name a Risk Owner (a role, not a person's private data), a Target Date and a Status. Review at least as often as the acceptance table in 02_Scales requires."),
        ("Step 9 — Report", "Open 05_Dashboard for counts by rating and 06_Risk_Matrix for the heat map. Use them in management reviews (ISO 27001 clause 9.3)."),
    ]
    for label, text in steps:
        line(label, text, height=34)

    section("4. Alignment with standards")
    line(
        "ISO/IEC 27001:2022",
        "Clause 6.1.2 (information security risk assessment): defined scales, consistent criteria, identified risks, owners, analysis and evaluation. "
        "Clause 6.1.3 (risk treatment): treatment decision, controls and Annex A mapping — usable as input to the Statement of Applicability. "
        "Clauses 8.2/8.3 (execute assessment and treatment) and 9.3 (management review: use the Dashboard).",
        height=62,
    )
    line(
        "NIST CSF 2.0",
        "GV.RM (risk management strategy: scales and acceptance criteria), ID.AM (asset management: application inventory), "
        "ID.RA (risk assessment: register), plus the Protect/Detect/Respond/Recover categories referenced in the Control Mapping column.",
        height=48,
    )
    line(
        "ISO/IEC 42001:2023",
        "Clause 6.1.2 (AI risk assessment), 6.1.3 (AI risk treatment) and 6.1.4 (AI system impact assessment). The 'Contains AI/ML?' flag in the inventory "
        "identifies applications that need AI-specific scenarios (data quality, bias, model drift, third-party AI services). Annex A references can be entered in Control Mapping.",
        height=62,
    )

    section("5. Workbook guide")
    for label, text in [
        (S_SCALES, "Likelihood and Impact scales, plus risk acceptance criteria by rating."),
        (S_INV, "Application inventory with business criticality, data classification and hosting."),
        (S_REG, "Main sheet: inherent and residual scoring, treatment, controls, owners and review dates (rows 5–50)."),
        (S_DASH, "Automatic counts by rating (inherent vs residual), KPIs, prioritisation guidance and a chart."),
        (S_MATRIX, "5×5 heat map with legend and live counts of risks per cell."),
    ]:
        line(label, text, height=20)

    section("6. Conventions")
    line("Blue text", "Input cells — type here.", label_font=f(bold=True, color=INPUT_BLUE), height=18)
    line("Black text, grey fill", "Formulas — do not overwrite (Application Name, Scores, Ratings).", height=18)
    line("Red date / red score", "Red target date = overdue and not closed. Red residual score = higher than inherent score (check your scoring).", height=32)

    section("7. Important notes")
    line(
        "Open source",
        "Released under the MIT License. Free to use, modify and redistribute — attribution to the original template is appreciated but not required beyond the licence text.",
        height=34,
    )
    line(
        "No personal data",
        "Use roles (e.g. 'Head of HR'), never personal names or contact details. Do not store credentials, client names or confidential incident details in this file.",
        height=34,
    )
    line(
        "Disclaimer",
        "This template supports — but does not replace — professional judgement, legal advice or certification audits. Scores are only as good as the evidence behind them.",
        height=34,
    )
    page_setup(ws, landscape=True)


# --------------------------------------------------------------------------- #
# 02_Scales
# --------------------------------------------------------------------------- #
def build_scales(wb: Workbook):
    ws = wb.create_sheet(S_SCALES)
    ws.sheet_properties.tabColor = MID_BLUE
    ws.sheet_view.showGridLines = False
    title(ws, "Scales and Acceptance Criteria", "Calibrate the example thresholds to your organisation before first use.", width_cols=6)
    widths(ws, {"A": 14, "B": 18, "C": 46, "D": 40, "E": 44, "F": 48})

    # ---- Likelihood
    put(ws, "A3", "LIKELIHOOD  (how likely is the scenario to occur within the next 12 months?)", font=f(bold=True, color=NAVY, size=11))
    header(ws, 4, ["Score", "Level", "Description", "Approx. frequency", "Indicative probability", "Example (application context)"])
    likelihood = [
        (1, "Rare", "May occur only in exceptional circumstances.", "Less than once in 10 years", "< 5 %",
         "Compromise of a hardened, patched, internal-only application through a previously unknown flaw."),
        (2, "Unlikely", "Could occur at some time but not expected.", "Once every 3–10 years", "5 – 20 %",
         "Misuse of a privileged account in an application that has strong logging and regular access reviews."),
        (3, "Possible", "Might occur; has happened in similar organisations.", "Once every 1–3 years", "20 – 50 %",
         "Exploitation of a known vulnerability on an internet-facing application patched on a monthly cycle."),
        (4, "Likely", "Will probably occur in most circumstances.", "About once a year", "50 – 80 %",
         "Credential-stuffing attempts against a customer portal that offers no multi-factor authentication."),
        (5, "Almost certain", "Expected to occur; already happening or imminent.", "Several times a year / ongoing", "> 80 %",
         "Phishing-led account takeover on a SaaS application with no MFA while active campaigns target the sector."),
    ]
    for i, rowdata in enumerate(likelihood):
        r = 5 + i
        for j, v in enumerate(rowdata):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.font = f(bold=(j < 2))
            c.alignment = Alignment(horizontal="center" if j in (0, 1, 4) else "left", vertical="top", wrap_text=True)
            c.border = BORDER
        ws.row_dimensions[r].height = 48

    # ---- Impact
    put(ws, "A11", "IMPACT  (what is the worst credible consequence?)  — thresholds are examples in a generic currency unit", font=f(bold=True, color=NAVY, size=11))
    header(ws, 12, ["Score", "Level", "Business / Financial", "Operational", "Data / Compliance / Reputation"])
    impact = [
        (1, "Negligible", "Direct and indirect loss under 5,000.", "Disruption under 1 hour; no workaround required.",
         "No personal or confidential data exposed. No regulatory interest. No external visibility."),
        (2, "Minor", "Loss of 5,000 – 50,000.", "Degraded service for up to 4 hours; easy workaround.",
         "Limited internal data exposed. No notification duty. Minor internal embarrassment."),
        (3, "Moderate", "Loss of 50,000 – 250,000; minor contract penalties.", "Outage of a non-critical process up to 1 business day.",
         "Confidential data of limited scope exposed. Possible regulator enquiry. Isolated customer complaints."),
        (4, "Major", "Loss of 250,000 – 1,000,000; loss of a significant customer.", "Outage of a critical process for 1–3 business days.",
         "Sensitive or personal data of many individuals. Mandatory breach notification. Regulatory penalties. Regional media coverage."),
        (5, "Severe", "Loss above 1,000,000 or a threat to business viability.", "Critical process unavailable for more than 3 days, or safety impact.",
         "Large-scale exposure of special-category data. Major sanctions or litigation. Sustained national media. Loss of licence to operate."),
    ]
    for i, rowdata in enumerate(impact):
        r = 13 + i
        for j, v in enumerate(rowdata):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.font = f(bold=(j < 2))
            c.alignment = Alignment(horizontal="center" if j in (0, 1) else "left", vertical="top", wrap_text=True)
            c.border = BORDER
        ws.row_dimensions[r].height = 62

    # ---- Acceptance criteria
    put(ws, "A19", "RISK ACCEPTANCE CRITERIA  (suggested — align with your risk appetite statement)", font=f(bold=True, color=NAVY, size=11))
    header(ws, 20, ["Rating", "Score range", "Required response", "Who may accept the risk", "Treatment deadline", "Review frequency"])
    accept = {
        "Low": ("Accept by default. Monitor for change.", "Application / risk owner", "None required", "Annually"),
        "Medium": ("Mitigate, or accept with a documented justification.", "Risk owner, informed to security function", "Within 6 months", "Every 6 months"),
        "High": ("Treatment plan mandatory with owner, budget and date. Acceptance is exceptional.", "Senior management (e.g. department head + CISO)", "Within 90 days", "Quarterly"),
        "Critical": ("Immediate escalation. Apply interim compensating controls now; treatment plan within 30 days.", "Executive management only, with written rationale", "Within 30 days (interim controls immediately)", "Monthly until below High"),
    }
    # list from highest to lowest for readability
    for i, (name, lo, hi, bg, fg) in enumerate(reversed(BANDS)):
        r = 21 + i
        vals = [name, f"{lo} – {hi}", *accept[name]]
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.border = BORDER
            c.alignment = Alignment(horizontal="center" if j < 2 else "left", vertical="top", wrap_text=True)
            c.font = f(bold=(j < 2), color=fg if j == 0 else "000000")
            if j == 0:
                c.fill = fill(bg)
        ws.row_dimensions[r].height = 48
    put(ws, "A26", "Note: rating thresholds (1–4 / 5–9 / 10–16 / 17–25) are built into the formulas of 04_Risk_Register and the Dashboard. "
        "If you change them, update the rating formulas in columns K and R and the legend on 06_Risk_Matrix.",
        font=f(italic=True, color="595959"))
    ws.merge_cells("A26:F26")
    ws.row_dimensions[26].height = 32
    page_setup(ws)


# --------------------------------------------------------------------------- #
# 03_Application_Inventory
# --------------------------------------------------------------------------- #
INV_HEADERS = [
    "App ID", "Application Name", "Business Owner (role)", "Technical Owner (role)",
    "Business Criticality", "Data Classification", "Hosting", "Users / Scope",
    "Internet Facing?", "Contains AI/ML?", "Last Risk Assessment", "Notes",
]

INV_SAMPLES = [
    ("APP-001", "Workforce HRIS", "Head of HR", "IT Applications Manager", "Critical", "Restricted", "SaaS",
     "All employees (~1,200)", "Yes", "No", date(2026, 9, 15),
     "Payroll, performance and personal data. SSO through the corporate identity provider."),
    ("APP-002", "Enterprise ERP (Finance & Supply Chain)", "Finance Director", "ERP Platform Lead", "Critical", "Confidential",
     "Private cloud", "Finance and procurement (~350 users)", "No", "No", date(2026, 8, 20),
     "Core financial ledger. Quarter-end availability is critical."),
    ("APP-003", "Customer Self-Service Portal", "Head of Customer Experience", "Digital Platform Lead", "High", "Confidential",
     "Public cloud (IaaS/PaaS)", "External customers (~40,000 accounts)", "Yes", "No", date(2026, 9, 30),
     "Order history, invoices and profile data. Peak load during seasonal campaigns."),
    ("APP-004", "AI Recommendation Engine", "Head of Data & Analytics", "ML Platform Lead", "High", "Confidential",
     "Public cloud (IaaS/PaaS)", "Feeds the customer portal; used by internal analysts", "No", "Yes", date(2026, 10, 10),
     "Uses a third-party model API for some features. Training data includes customer behaviour data."),
    ("APP-005", "CRM Platform", "Sales Director", "CRM Administrator", "High", "Confidential",
     "SaaS", "Sales and support (~180 users)", "Yes", "Yes", None,
     "Embedded lead-scoring feature. Not yet assessed — add risks to the register."),
]

CRIT = ["Critical", "High", "Medium", "Low"]
CLASSIF = ["Public", "Internal", "Confidential", "Restricted"]
HOSTING = ["On-premises", "Private cloud", "Public cloud (IaaS/PaaS)", "SaaS", "Hybrid"]
YESNO = ["Yes", "No"]


def list_dv(options, *, warn=False):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    dv.showErrorMessage = True
    dv.errorStyle = "warning" if warn else "stop"
    dv.errorTitle = "Invalid entry"
    dv.error = "Please choose a value from the list."
    return dv


def build_inventory(wb: Workbook):
    ws = wb.create_sheet(S_INV)
    ws.sheet_properties.tabColor = "548235"
    title(ws, "Application Inventory", "One row per application. Blue = input. Sample rows are fictional — overwrite or delete them.", width_cols=12)
    widths(ws, {"A": 11, "B": 32, "C": 24, "D": 24, "E": 14, "F": 15, "G": 22, "H": 34, "I": 11, "J": 11, "K": 14, "L": 52})
    header(ws, 4, INV_HEADERS, height=40)

    for r in range(INV_FIRST, INV_LAST + 1):
        idx = r - INV_FIRST
        sample = INV_SAMPLES[idx] if idx < len(INV_SAMPLES) else None
        for ci in range(1, 13):
            c = ws.cell(row=r, column=ci)
            if sample and sample[ci - 1] is not None:
                c.value = sample[ci - 1]
            c.font = f(color=INPUT_BLUE, bold=(ci == 1))
            c.border = BORDER
            c.alignment = Alignment(
                vertical="top", wrap_text=True,
                horizontal="center" if ci in (1, 5, 6, 9, 10, 11) else "left",
            )
        ws.cell(row=r, column=11).number_format = "yyyy-mm-dd"
        if sample:
            ws.row_dimensions[r].height = 44

    rng = lambda col: f"{col}{INV_FIRST}:{col}{INV_LAST}"
    for col, opts in (("E", CRIT), ("F", CLASSIF), ("G", HOSTING), ("I", YESNO), ("J", YESNO)):
        dv = list_dv(opts)
        ws.add_data_validation(dv)
        dv.add(rng(col))
    dv_date = DataValidation(type="date", operator="greaterThanOrEqual", formula1="36526", allow_blank=True)
    dv_date.errorStyle = "warning"
    dv_date.showErrorMessage = True
    dv_date.error = "Enter a valid date (yyyy-mm-dd)."
    ws.add_data_validation(dv_date)
    dv_date.add(rng("K"))

    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:L{INV_LAST}"
    page_setup(ws)


# --------------------------------------------------------------------------- #
# 04_Risk_Register
# --------------------------------------------------------------------------- #
REG_HEADERS = [
    "Risk ID", "Application ID", "Application Name", "Risk Scenario / Description", "Threat",
    "Vulnerability / Weakness", "Asset / Data Affected",
    "Inherent Likelihood (1-5)", "Inherent Impact (1-5)", "Inherent Score", "Inherent Rating",
    "Treatment Decision", "Existing / Planned Controls", "Control Mapping (ISO 27001 Annex A / NIST CSF)",
    "Residual Likelihood (1-5)", "Residual Impact (1-5)", "Residual Score", "Residual Rating",
    "Risk Owner (role)", "Target Date", "Status", "Last Review", "Comments",
]

STATUS = ["Open", "In progress", "Mitigated", "Accepted", "Closed"]
TREATMENT = ["Mitigate", "Accept", "Transfer", "Avoid"]

# (id, app_id, scenario, threat, vulnerability, asset, iL, iI, treatment, controls, mapping,
#  rL, rI, owner, target, status, last_review, comments)
REG_SAMPLES = [
    ("R-001", "APP-001",
     "Takeover of an HR administrator account gives an attacker access to employee personal, payroll and banking data.",
     "Phishing / credential stuffing by an external attacker",
     "Phishing-resistant MFA not enforced for the HR administrator role; no alerting on bulk exports.",
     "Employee PII, payroll and banking details",
     4, 5, "Mitigate",
     "Planned: enforce phishing-resistant MFA for privileged roles; conditional access; SIEM alert on bulk export; quarterly access review.",
     "ISO 27001: A.5.15, A.5.17, A.8.2, A.8.5, A.8.15 | NIST CSF 2.0: PR.AA, DE.CM",
     2, 4, "Head of HR", date(2026, 12, 31), "In progress", date(2026, 10, 15), "MFA rollout about 60% complete."),
    ("R-002", "APP-002",
     "Delayed patching of the ERP application and database tiers lets an attacker exploit a known vulnerability, causing an outage or tampering with financial data.",
     "Exploitation of known vulnerabilities (ransomware or fraud actor)",
     "Average patch cycle above 90 days; limited test environment slows deployment.",
     "Financial records; ERP availability",
     3, 5, "Mitigate",
     "Risk-based patch SLA (critical within 14 days); refresh test environment; authenticated vulnerability scanning; immutable backups.",
     "ISO 27001: A.8.8, A.8.13, A.8.32 | NIST CSF 2.0: PR.PS, ID.RA",
     2, 4, "Finance Director", date(2027, 1, 31), "Open", date(2026, 10, 2), "Business agreed a monthly patch window."),
    ("R-003", "APP-003",
     "Automated credential-stuffing against the customer portal leads to account takeover and fraudulent orders.",
     "Automated credential-stuffing bots",
     "No MFA option for customers; weak rate limiting; no bot detection.",
     "Customer accounts, personal data, order history",
     5, 3, "Mitigate",
     "Rate limiting; bot management / WAF rules; optional MFA; breached-password screening; login-anomaly alerts.",
     "ISO 27001: A.5.34, A.8.5, A.8.16, A.8.20 | NIST CSF 2.0: PR.AA, DE.CM",
     3, 3, "Head of Customer Experience", date(2026, 12, 15), "In progress", date(2026, 10, 20), "WAF rules in test; MFA design pending."),
    ("R-004", "APP-004",
     "Poisoned or biased training data makes the AI recommendation engine produce harmful or discriminatory outputs, creating regulatory and reputational damage.",
     "Data poisoning / insufficient data governance",
     "No documented data-provenance checks; no bias or robustness testing before release; limited model monitoring.",
     "Training datasets; model integrity; customer trust",
     3, 4, "Mitigate",
     "Data-provenance and quality gates; bias and robustness testing; model monitoring with rollback; AI system impact assessment.",
     "ISO 42001: Annex A.5, A.6, A.7 | ISO 27001: A.5.9, A.8.32 | NIST CSF 2.0: GV.RM, ID.RA",
     2, 3, "Head of Data & Analytics", date(2027, 3, 31), "In progress", date(2026, 10, 10), "AI impact assessment scheduled for next quarter."),
    ("R-005", "APP-004",
     "Customer personal data leaks through prompts, logs or a third-party model API used by the AI engine.",
     "Data leakage through a third-party AI service",
     "Prompts and outputs logged in clear text; vendor data-retention terms not reviewed.",
     "Customer personal data; proprietary model inputs",
     3, 4, "Mitigate",
     "Redact or pseudonymise data before sending; zero-retention vendor terms; minimise logging; DLP on outbound API traffic.",
     "ISO 27001: A.5.19, A.5.20, A.5.34, A.8.12 | ISO 42001: Annex A.10 | NIST CSF 2.0: GV.SC, PR.DS",
     2, 3, "CISO", date(2026, 12, 1), "Open", date(2026, 10, 10), "Vendor review in progress."),
    ("R-006", "APP-002",
     "A retired reporting add-on remains installed on the ERP server and offers a small additional attack surface.",
     "Opportunistic exploitation of unused components",
     "Component not removed after migration; isolated but not decommissioned.",
     "ERP server configuration",
     2, 2, "Accept",
     "Network-isolated with no external access; reviewed annually; removal scheduled with next major upgrade.",
     "ISO 27001: A.8.9 | NIST CSF 2.0: PR.PS",
     2, 2, "ERP Platform Lead", date(2027, 6, 30), "Accepted", date(2026, 9, 5), "Accepted by risk owner; revisit at upgrade."),
    ("R-007", "APP-003",
     "A prolonged outage of the primary cloud region makes the customer portal unavailable during a peak sales period.",
     "Cloud provider regional outage",
     "Single-region deployment; disaster-recovery plan not tested in the last 12 months.",
     "Portal availability; revenue during peak periods",
     2, 4, "Transfer",
     "Multi-zone deployment; tested DR plan; business-interruption insurance covering cloud dependency.",
     "ISO 27001: A.5.29, A.5.30, A.8.14 | NIST CSF 2.0: PR.IR, RC.RP",
     1, 4, "Digital Platform Lead", date(2026, 11, 30), "In progress", date(2026, 9, 25), "Insurance renewal in November."),
]

COLS = {name: get_column_letter(i + 1) for i, name in enumerate([
    "id", "app", "appname", "scenario", "threat", "vuln", "asset", "iL", "iI", "iS", "iR",
    "treat", "controls", "mapping", "rL", "rI", "rS", "rR", "owner", "target", "status", "review", "comments",
])}


def rating_formula(score_cell: str) -> str:
    return (f'=IF({score_cell}="","",IF({score_cell}>=17,"Critical",IF({score_cell}>=10,"High",'
            f'IF({score_cell}>=5,"Medium","Low"))))')


def build_register(wb: Workbook):
    ws = wb.create_sheet(S_REG)
    ws.sheet_properties.tabColor = "C55A11"
    title(ws, "Risk Register", "Likelihood × Impact. Blue = input; grey cells are formulas. Rows 5–50 are ready to use. Sample risks are fictional.", width_cols=23)

    # group banner (row 3)
    groups = [("RISK IDENTIFICATION", "A", "G", "2F5597"), ("INHERENT RISK", "H", "K", "C55A11"),
              ("TREATMENT", "L", "N", "548235"), ("RESIDUAL RISK", "O", "R", "7030A0"), ("GOVERNANCE", "S", "W", "595959")]
    for label, c1, c2, color in groups:
        ws.merge_cells(f"{c1}3:{c2}3")
        cell = ws[f"{c1}3"]
        cell.value = label
        cell.font = f(bold=True, color="FFFFFF")
        cell.fill = fill(color)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 20

    header(ws, 4, REG_HEADERS, height=52)
    widths(ws, {
        "A": 9, "B": 12, "C": 26, "D": 46, "E": 28, "F": 38, "G": 28, "H": 11, "I": 11, "J": 10, "K": 11,
        "L": 13, "M": 44, "N": 40, "O": 11, "P": 11, "Q": 10, "R": 11, "S": 22, "T": 12, "U": 12, "V": 12, "W": 34,
    })

    C = COLS
    formula_cols = {C["appname"], C["iS"], C["iR"], C["rS"], C["rR"]}
    center_cols = {C["id"], C["app"], C["iL"], C["iI"], C["iS"], C["iR"], C["treat"], C["rL"], C["rI"], C["rS"], C["rR"],
                   C["target"], C["status"], C["review"]}

    for r in range(REG_FIRST, REG_LAST + 1):
        sample = REG_SAMPLES[r - REG_FIRST] if r - REG_FIRST < len(REG_SAMPLES) else None
        values = {}
        if sample:
            (rid, app, scenario, threat, vuln, asset, iL, iI, treat, controls, mapping,
             rL, rI, owner, target, status, review, comments) = sample
            values = {C["id"]: rid, C["app"]: app, C["scenario"]: scenario, C["threat"]: threat, C["vuln"]: vuln,
                      C["asset"]: asset, C["iL"]: iL, C["iI"]: iI, C["treat"]: treat, C["controls"]: controls,
                      C["mapping"]: mapping, C["rL"]: rL, C["rI"]: rI, C["owner"]: owner, C["target"]: target,
                      C["status"]: status, C["review"]: review, C["comments"]: comments}
        # formulas
        values[C["appname"]] = (
            f'=IF(${C["app"]}{r}="","",IFERROR(INDEX(\'{S_INV}\'!$B${INV_FIRST}:$B${INV_LAST},'
            f'MATCH(${C["app"]}{r},\'{S_INV}\'!$A${INV_FIRST}:$A${INV_LAST},0)),"ID not found in inventory"))'
        )
        values[C["iS"]] = f'=IF(AND(ISNUMBER({C["iL"]}{r}),ISNUMBER({C["iI"]}{r})),{C["iL"]}{r}*{C["iI"]}{r},"")'
        values[C["iR"]] = rating_formula(f'{C["iS"]}{r}')
        values[C["rS"]] = f'=IF(AND(ISNUMBER({C["rL"]}{r}),ISNUMBER({C["rI"]}{r})),{C["rL"]}{r}*{C["rI"]}{r},"")'
        values[C["rR"]] = rating_formula(f'{C["rS"]}{r}')

        for col_letter in [get_column_letter(i) for i in range(1, 24)]:
            c = ws[f"{col_letter}{r}"]
            if col_letter in values:
                c.value = values[col_letter]
            is_formula = col_letter in formula_cols
            c.font = f(color="000000" if is_formula else INPUT_BLUE, bold=(col_letter in {C["iS"], C["iR"], C["rS"], C["rR"]}))
            if is_formula:
                c.fill = fill(PALE)
            c.border = BORDER
            c.alignment = Alignment(
                vertical="top", wrap_text=True,
                horizontal="center" if col_letter in center_cols else "left",
            )
        ws[f"{C['target']}{r}"].number_format = "yyyy-mm-dd"
        ws[f"{C['review']}{r}"].number_format = "yyyy-mm-dd"
        if sample:
            ws.row_dimensions[r].height = 92

    first, last = REG_FIRST, REG_LAST

    # ---- data validations
    dv_app = DataValidation(type="list", formula1=f"='{S_INV}'!$A${INV_FIRST}:$A${INV_LAST}", allow_blank=True)
    dv_app.showErrorMessage = True
    dv_app.errorStyle = "warning"
    dv_app.errorTitle = "Unknown application"
    dv_app.error = "This ID is not in the Application Inventory. Add the application there first."
    ws.add_data_validation(dv_app)
    dv_app.add(f"{C['app']}{first}:{C['app']}{last}")

    dv_score = list_dv(["1", "2", "3", "4", "5"])
    dv_score.errorTitle = "Score must be 1–5"
    dv_score.error = "Enter a whole number from 1 to 5 (see 02_Scales)."
    ws.add_data_validation(dv_score)
    for key in ("iL", "iI", "rL", "rI"):
        dv_score.add(f"{C[key]}{first}:{C[key]}{last}")

    dv_treat = list_dv(TREATMENT)
    ws.add_data_validation(dv_treat)
    dv_treat.add(f"{C['treat']}{first}:{C['treat']}{last}")

    dv_status = list_dv(STATUS)
    ws.add_data_validation(dv_status)
    dv_status.add(f"{C['status']}{first}:{C['status']}{last}")

    dv_date = DataValidation(type="date", operator="greaterThanOrEqual", formula1="36526", allow_blank=True)
    dv_date.errorStyle = "warning"
    dv_date.showErrorMessage = True
    dv_date.error = "Enter a valid date (yyyy-mm-dd)."
    ws.add_data_validation(dv_date)
    dv_date.add(f"{C['target']}{first}:{C['target']}{last}")
    dv_date.add(f"{C['review']}{first}:{C['review']}{last}")

    # ---- conditional formatting (ratings + scores)
    def band_rules(col, text_based):
        rng = f"{col}{first}:{col}{last}"
        for name, lo, hi, bg, fg in BANDS:
            if text_based:
                formula = f'${col}{first}="{name}"'
            else:
                formula = f"AND(ISNUMBER(${col}{first}),${col}{first}>={lo},${col}{first}<={hi})"
            ws.conditional_formatting.add(
                rng, FormulaRule(formula=[formula], fill=fill(bg), font=Font(name=FONT, bold=True, color=fg), stopIfTrue=True))

    for key in ("iR", "rR"):
        band_rules(C[key], True)
    for key in ("iS", "rS"):
        band_rules(C[key], False)

    # residual higher than inherent -> red bold (scoring sanity check)
    ws.conditional_formatting.add(
        f"{C['rS']}{first}:{C['rS']}{last}",
        FormulaRule(formula=[f"AND(ISNUMBER({C['rS']}{first}),ISNUMBER({C['iS']}{first}),{C['rS']}{first}>{C['iS']}{first})"],
                    font=Font(name=FONT, bold=True, color="FF0000"), fill=fill("FFD9D9")))
    # overdue target date
    ws.conditional_formatting.add(
        f"{C['target']}{first}:{C['target']}{last}",
        FormulaRule(formula=[f'AND(ISNUMBER({C["target"]}{first}),{C["target"]}{first}<TODAY(),'
                             f'{C["status"]}{first}<>"Closed",{C["status"]}{first}<>"Mitigated",{C["status"]}{first}<>"Accepted")'],
                    font=Font(name=FONT, bold=True, color="FF0000")))

    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:W{last}"
    page_setup(ws)
    ws.print_title_rows = "3:4"


# --------------------------------------------------------------------------- #
# 05_Dashboard
# --------------------------------------------------------------------------- #
def build_dashboard(wb: Workbook):
    ws = wb.create_sheet(S_DASH)
    ws.sheet_properties.tabColor = "7030A0"
    ws.sheet_view.showGridLines = False
    title(ws, "Risk Dashboard", f"Figures update automatically from {S_REG}.", width_cols=14)
    widths(ws, {"A": 26, "B": 14, "C": 14, "D": 14, "E": 14, "F": 4, "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14, "M": 14, "N": 14})

    R = f"'{S_REG}'!"
    rng = lambda col: f"{R}${col}${REG_FIRST}:${col}${REG_LAST}"
    C = COLS

    # ---- KPIs (row 4 labels / row 5 values)
    kpis = [
        ("Risks recorded", f"=COUNTA({rng(C['id'])})", "0"),
        ("Open (not Closed)", f'=COUNTA({rng(C["id"])})-COUNTIF({rng(C["status"])},"Closed")', "0"),
        ("Highest residual score", f'=IF(COUNT({rng(C["rS"])})=0,"–",MAX({rng(C["rS"])}))', "0"),
        ("Average inherent score", f'=IF(COUNT({rng(C["iS"])})=0,"–",AVERAGE({rng(C["iS"])}))', "0.0"),
        ("Average residual score", f'=IF(COUNT({rng(C["rS"])})=0,"–",AVERAGE({rng(C["rS"])}))', "0.0"),
        ("Overdue actions", f'=COUNTIFS({rng(C["target"])},"<"&TODAY(),{rng(C["status"])},"<>Closed",{rng(C["status"])},"<>Mitigated",{rng(C["status"])},"<>Accepted")', "0"),
    ]
    cols = ["A", "B", "C", "D", "E"]
    # lay KPIs in a 2-row x 3 grid starting at A4 using merged pairs
    positions = [("A", "A"), ("B", "C"), ("D", "E"), ("G", "H"), ("I", "J"), ("K", "L")]
    for (label, formula, fmt), (c1, c2) in zip(kpis, positions):
        if c1 != c2:
            ws.merge_cells(f"{c1}4:{c2}4")
            ws.merge_cells(f"{c1}5:{c2}5")
        put(ws, f"{c1}4", label, font=f(bold=True, color="FFFFFF"), bg=MID_BLUE,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        put(ws, f"{c1}5", formula, font=f(bold=True, size=18, color=NAVY), bg=LIGHT_BLUE,
            align=Alignment(horizontal="center", vertical="center"), fmt=fmt)
        for cc in {c1, c2}:
            ws[f"{cc}4"].fill = fill(MID_BLUE)
            ws[f"{cc}5"].fill = fill(LIGHT_BLUE)
    ws.row_dimensions[4].height = 30
    ws.row_dimensions[5].height = 40

    # ---- Rating table (inherent vs residual)
    put(ws, "A7", "RISKS BY RATING", font=f(bold=True, color=NAVY, size=11))
    header(ws, 8, ["Rating", "Score range", "Inherent", "Residual", "Reduction"], height=28)
    order = list(reversed(BANDS))
    for i, (name, lo, hi, bg, fg) in enumerate(order):
        r = 9 + i
        put(ws, f"A{r}", name, font=f(bold=True, color=fg), bg=bg, align=Alignment(horizontal="center", vertical="center"), border=True)
        put(ws, f"B{r}", f"{lo} – {hi}", align=Alignment(horizontal="center", vertical="center"), border=True)
        put(ws, f"C{r}", f'=COUNTIF({rng(C["iR"])},$A{r})', font=f(bold=True), align=Alignment(horizontal="center", vertical="center"), border=True)
        put(ws, f"D{r}", f'=COUNTIF({rng(C["rR"])},$A{r})', font=f(bold=True), align=Alignment(horizontal="center", vertical="center"), border=True)
        put(ws, f"E{r}", f"=C{r}-D{r}", align=Alignment(horizontal="center", vertical="center"), border=True, fmt="+0;-0;0")
        ws.row_dimensions[r].height = 22
    put(ws, "A13", "Total rated", font=f(bold=True), bg=PALE, align=Alignment(horizontal="center", vertical="center"), border=True)
    put(ws, "B13", "", bg=PALE, border=True)
    put(ws, "C13", "=SUM(C9:C12)", font=f(bold=True), bg=PALE, align=Alignment(horizontal="center", vertical="center"), border=True)
    put(ws, "D13", "=SUM(D9:D12)", font=f(bold=True), bg=PALE, align=Alignment(horizontal="center", vertical="center"), border=True)
    put(ws, "E13", "", bg=PALE, border=True)

    # ---- Treatment & status tables
    put(ws, "A15", "BY TREATMENT DECISION", font=f(bold=True, color=NAVY, size=11))
    header(ws, 16, ["Decision", "Count"], height=24)
    for i, t in enumerate(TREATMENT):
        r = 17 + i
        put(ws, f"A{r}", t, align=Alignment(horizontal="left", vertical="center"), border=True)
        put(ws, f"B{r}", f'=COUNTIF({rng(C["treat"])},$A{r})', font=f(bold=True), align=Alignment(horizontal="center", vertical="center"), border=True)
    put(ws, "A22", "BY STATUS", font=f(bold=True, color=NAVY, size=11))
    header(ws, 23, ["Status", "Count"], height=24)
    for i, s in enumerate(STATUS):
        r = 24 + i
        put(ws, f"A{r}", s, align=Alignment(horizontal="left", vertical="center"), border=True)
        put(ws, f"B{r}", f'=COUNTIF({rng(C["status"])},$A{r})', font=f(bold=True), align=Alignment(horizontal="center", vertical="center"), border=True)

    # ---- Chart (inherent vs residual by rating)
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Inherent vs residual risks by rating"
    chart.y_axis.title = "Number of risks"
    chart.y_axis.majorGridlines = None
    data = Reference(ws, min_col=3, max_col=4, min_row=8, max_row=12)
    cats = Reference(ws, min_col=1, min_row=9, max_row=12)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8.5
    chart.width = 17
    chart.series[0].graphicalProperties.solidFill = "C55A11"
    chart.series[1].graphicalProperties.solidFill = "2F5597"
    ws.add_chart(chart, "G7")

    # ---- Prioritisation guidance
    put(ws, "A31", "HOW TO PRIORITISE", font=f(bold=True, color=NAVY, size=11))
    tips = [
        "1. Work top-down: Critical first, then High. Within a band, start with the highest residual score.",
        "2. Treat a risk on a Critical business application (see inventory) before an equal-scored risk on a Low-criticality one.",
        "3. Prefer quick wins that cut several risks at once (e.g. MFA, patch SLA, logging) over one-off fixes.",
        "4. Residual should be lower than inherent. If it is not, the planned controls are not adding value — revisit them.",
        "5. Never leave Critical or High risks without an owner, a target date and an interim control.",
        "6. Re-score after every significant change (new feature, new supplier, incident, audit finding).",
    ]
    for i, t in enumerate(tips):
        r = 32 + i
        ws.merge_cells(f"A{r}:L{r}")
        put(ws, f"A{r}", t, align=Alignment(vertical="center", wrap_text=True))
        ws.row_dimensions[r].height = 20

    put(ws, "A39", "ACTIONS BY RATING (suggested)", font=f(bold=True, color=NAVY, size=11))
    header(ws, 40, ["Rating", "Required action"], height=24)
    ws.merge_cells("B40:L40")
    actions = {
        "Critical": "Escalate to executive management immediately. Apply interim controls now. Treatment plan within 30 days. Review monthly.",
        "High": "Treatment plan with owner, budget and date within 90 days. Acceptance only by senior management. Review quarterly.",
        "Medium": "Mitigate, or accept with documented justification by the risk owner. Review every 6 months.",
        "Low": "Accept by default and monitor. Review annually.",
    }
    for i, (name, lo, hi, bg, fg) in enumerate(order):
        r = 41 + i
        put(ws, f"A{r}", name, font=f(bold=True, color=fg), bg=bg, align=Alignment(horizontal="center", vertical="center"), border=True)
        ws.merge_cells(f"B{r}:L{r}")
        put(ws, f"B{r}", actions[name], align=Alignment(vertical="center", wrap_text=True), border=True)
        ws.row_dimensions[r].height = 24
    page_setup(ws)


# --------------------------------------------------------------------------- #
# 06_Risk_Matrix
# --------------------------------------------------------------------------- #
LIK_NAMES = {1: "Rare", 2: "Unlikely", 3: "Possible", 4: "Likely", 5: "Almost certain"}
IMP_NAMES = {1: "Negligible", 2: "Minor", 3: "Moderate", 4: "Major", 5: "Severe"}


def _matrix(ws, top, caption, *, kind):
    """Draw a 5x5 grid whose first header row is `top`.
    kind: 'score' -> L*I formula, 'inherent' / 'residual' -> COUNTIFS from the register."""
    C = COLS
    put(ws, f"A{top - 1}", caption, font=f(bold=True, color=NAVY, size=11))
    # impact names and numbers
    put(ws, f"A{top}", "", bg=None)
    ws.merge_cells(f"C{top - 1}:G{top - 1}")
    put(ws, f"C{top - 1}", "IMPACT →", font=f(bold=True, color="FFFFFF"), bg=NAVY, align=Alignment(horizontal="center", vertical="center"))
    for i in range(1, 6):
        col = get_column_letter(2 + i)
        put(ws, f"{col}{top}", f"{i} · {IMP_NAMES[i]}", font=f(bold=True, color="FFFFFF"), bg=MID_BLUE,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True), border=True)
        put(ws, f"{col}{top + 1}", i, font=f(bold=True, color="595959"), align=Alignment(horizontal="center", vertical="center"), fmt="0")
    ws.row_dimensions[top].height = 32
    ws.row_dimensions[top + 1].height = 16
    for k, L in enumerate(range(5, 0, -1)):
        r = top + 2 + k
        put(ws, f"A{r}", f"{L} · {LIK_NAMES[L]}", font=f(bold=True, color="FFFFFF"), bg=MID_BLUE,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True), border=True)
        put(ws, f"B{r}", L, font=f(bold=True, color="595959"), align=Alignment(horizontal="center", vertical="center"), fmt="0")
        ws.row_dimensions[r].height = 34
        for i in range(1, 6):
            col = get_column_letter(2 + i)
            name, bg, fg = band_for(L * i)
            if kind == "score":
                value = f"=$B{r}*{col}${top + 1}"
                fmt = "0"
            else:
                key_l, key_i = ("iL", "iI") if kind == "inherent" else ("rL", "rI")
                value = (f"=COUNTIFS('{S_REG}'!${C[key_l]}${REG_FIRST}:${C[key_l]}${REG_LAST},$B{r},"
                         f"'{S_REG}'!${C[key_i]}${REG_FIRST}:${C[key_i]}${REG_LAST},{col}${top + 1})")
                fmt = "0;-0;;@"   # hide zeros
            put(ws, f"{col}{r}", value, font=f(bold=True, color=fg, size=14), bg=bg,
                align=Alignment(horizontal="center", vertical="center"), border=True, fmt=fmt)
    # side label
    ws.merge_cells(f"A{top + 7}:B{top + 7}")
    put(ws, f"A{top + 7}", "↑ LIKELIHOOD", font=f(bold=True, color=NAVY), align=Alignment(horizontal="center", vertical="center"))


def build_matrix(wb: Workbook):
    ws = wb.create_sheet(S_MATRIX)
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.showGridLines = False
    title(ws, "Risk Matrix (5×5)", "Likelihood in rows, Impact in columns. Score = Likelihood × Impact.", width_cols=11)
    widths(ws, {"A": 20, "B": 6, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 4, "I": 14, "J": 50, "K": 4})

    _matrix(ws, 5, "SCORE MATRIX", kind="score")

    # legend (right of the first matrix)
    put(ws, "I4", "LEGEND", font=f(bold=True, color=NAVY, size=11))
    header(ws, 5, ["Rating", "Score range & meaning"], start_col=9, height=32)
    meanings = {
        "Low": "1–4  ·  Accept and monitor",
        "Medium": "5–9  ·  Mitigate or accept with justification",
        "High": "10–16  ·  Treatment plan within 90 days",
        "Critical": "17–25  ·  Escalate now; plan within 30 days",
    }
    for i, (name, lo, hi, bg, fg) in enumerate(reversed(BANDS)):
        r = 7 + i
        put(ws, f"I{r}", name, font=f(bold=True, color=fg), bg=bg, align=Alignment(horizontal="center", vertical="center"), border=True)
        put(ws, f"J{r}", meanings[name], align=Alignment(vertical="center", wrap_text=True), border=True)
    # row 6 of the legend block is the header spacer: keep clean
    _matrix(ws, 17, "INHERENT RISKS PER CELL (live count from the register)", kind="inherent")
    _matrix(ws, 29, "RESIDUAL RISKS PER CELL (live count from the register)", kind="residual")
    page_setup(ws, landscape=True)


# --------------------------------------------------------------------------- #
def main(out: Path):
    wb = Workbook()
    build_methodology(wb)
    build_scales(wb)
    build_inventory(wb)
    build_register(wb)
    build_dashboard(wb)
    build_matrix(wb)

    wb.properties.title = "Application Risk Assessment Template"
    wb.properties.subject = "Likelihood x Impact application risk assessment (ISO 27001, NIST CSF 2.0, ISO 42001)"
    wb.properties.creator = "Open-source contributors"
    wb.properties.description = "MIT-licensed template. All sample data is fictional."
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    default = Path(__file__).resolve().parent.parent / "Application-Risk-Assessment-Template.xlsx"
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else default)
