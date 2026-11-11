import os
import io
import re
import json
import zipfile
import tempfile
from datetime import datetime, timedelta, timezone
from functools import wraps
import jwt as pyjwt
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify, Response, make_response
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import mm
from num2words import num2words

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
    static_url_path="/static"
)

# ─── JWT Auth Configuration ───────────────────────────────────────────────────
JWT_SECRET     = os.environ.get("SECRET_KEY", "yayath-payroll-super-secret-key-2026-XkQ9")
JWT_ALGORITHM  = "HS256"
JWT_EXPIRY_HRS = 8          # Token valid for 8 hours
JWT_COOKIE     = "yb_auth"  # HttpOnly cookie name

AUTH_USER = os.environ.get("AUTH_USER", "HR123")
AUTH_PASS = os.environ.get("AUTH_PASS", "HR123#")
IS_VERCEL = os.environ.get("VERCEL", "0") == "1"

def _make_jwt(user_id: str) -> str:
    """Create a signed JWT token that expires in JWT_EXPIRY_HRS hours."""
    payload = {
        "sub": user_id,
        "iat": datetime.now(timezone.utc),
        "exp": datetime.now(timezone.utc) + timedelta(hours=JWT_EXPIRY_HRS),
    }
    return pyjwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)

def _verify_jwt(token: str):
    """Verify token and return payload dict, or None if invalid/expired."""
    try:
        return pyjwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
    except pyjwt.ExpiredSignatureError:
        return None
    except pyjwt.InvalidTokenError:
        return None

def login_required(f):
    """Decorator: protects a route with JWT cookie auth."""
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.cookies.get(JWT_COOKIE)
        if not token or not _verify_jwt(token):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def get_current_user() -> str | None:
    """Return logged-in user ID from JWT cookie, or None."""
    token = request.cookies.get(JWT_COOKIE)
    if not token:
        return None
    payload = _verify_jwt(token)
    return payload.get("sub") if payload else None

@app.context_processor
def inject_auth():
    """Make current_user available in every template automatically."""
    return {"current_user": get_current_user()}

UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), "uploads")
try:
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
except Exception:
    pass
ALLOWED = {".xlsx", ".xls", ".csv"}

# Comprehensive field specifications with fuzzy matching aliases
FIELDS = [
    # ── Identity ──────────────────────────────────────────────────────────────
    ("name",        "Employee Name",         ["empname", "emp name", "employee name", "employee_name", "name", "full name", "staff name", "worker name"]),
    ("code",        "Employee Code",         ["empcode", "emp code", "employee code", "employee_code", "emp_code", "code", "employee id", "emp id", "id", "staff no", "badge no"]),
    ("designation", "Designation",           ["designation", "role", "position", "title", "post", "job title"]),
    ("department",  "Department",            ["department", "dept", "department name", "dept name", "section"]),
    ("pan",         "PAN",                   ["pan", "pan number", "pan_no", "pan card", "panno", "pan no."]),
    ("joining_date","Joining Date",          ["joining date", "joining_date", "date of joining", "doj", "date_of_joining"]),
    ("location",    "Location",              ["location", "city", "work location", "site", "project site", "branch location"]),
    ("unit",        "Unit",                  ["unit", "branch", "business unit", "division"]),
    # ── Attendance ────────────────────────────────────────────────────────────
    ("total_days",  "Total Days",            ["totaldays", "total days", "total_days", "working days in month", "calendar days"]),
    ("paid_days",   "Paid Days",             ["paiddays", "paid days", "paid_days", "payable days", "net paid days"]),
    ("lwp",         "Leave Without Pay",     ["leave without pay", "lwp", "lop", "loss of pay", "unpaid", "absent", "leave_without_pay"]),
    ("present",     "Present Days",          ["present", "present days", "days worked", "total present"]),
    ("nh",          "NH",                    ["nh", "national holiday", "holidays", "festival holiday"]),
    ("woff",        "WOFF",                  ["woff", "weekly off", "weekly_off", "off days", "wo", "week off"]),
    # ── Employee IDs ─────────────────────────────────────────────────────────
    ("bank",        "Bank Account No.",      ["bank account no", "bank account", "bank_account", "account number", "a/c no", "account_no", "bank a/c", "acc no"]),
    ("pf_no",       "PF No.",               ["pf_no", "pf no", "pf number", "uan", "pf uan", "provident fund no", "uan no", "pf a/c"]),
    ("esi_no",      "ESI No.",              ["esi no", "esi no.", "esi number", "esi_no", "insurance no", "esic no", "ip number", "esic"]),
    ("leave_bal",   "Leave Balance",         ["leave bal", "leave balance", "leave_bal", "leaves", "balance leaves", "cl", "pl", "bal leaves"]),
    ("cost_center", "Cost Center",           ["cost center", "cost_center", "project", "cc"]),
    # ── Earnings ─────────────────────────────────────────────────────────────
    ("basic_entitled",  "Basic Entitled",    ["basic da", "basic +da", "basic+da", "basic", "basic salary", "entitled basic", "basic pay", "basic rate", "basic entitled"]),
    ("basic_earned",    "Basic Earned",      ["basic earned", "earned basic", "basic pay earned"]),
    ("basic_arrears",   "Basic Arrears",     ["basic arrears", "arrears basic", "basic arrear"]),
    ("hra_entitled",    "HRA Entitled",      ["hra", "house rent allowance", "hra rate", "hra entitled"]),
    ("hra_earned",      "HRA Earned",        ["hra earned", "earned hra"]),
    ("hra_arrears",     "HRA Arrears",       ["hra arrears", "arrears hra"]),
    ("add_allowance",   "Additional Allowance", ["additional allowance", "additional_allowance", "add allowance", "addl allowance", "allowance", "other allowance", "special allowance"]),
    ("bonus_entitled",  "Bonus Entitled",    ["bonus", "statutory bonus", "bonus entitled", "bonus rate", "incentive"]),
    ("bonus_earned",    "Bonus Earned",      ["bonus earned", "earned bonus"]),
    ("bonus_arrears",   "Bonus Arrears",     ["bonus arrears", "arrears bonus"]),
    # ── Deductions ───────────────────────────────────────────────────────────
    ("pf",          "PF Deduction",          ["pf", "pf ", "pf deduction", "pf amount", "provident fund", "epf", "pf ee", "pf employee"]),
    ("esi",         "ESI Deduction",         ["esi basic da", "esi basic", "esi", "esi deduction", "esi amount", "esic ee", "esi employee"]),
    ("prof_tax",    "Professional Tax",      ["proftax", "prof tax", "professional tax", "pt", "p tax", "ptax"]),
    # ── Totals (optional) ────────────────────────────────────────────────────
    ("gross_earning",   "Gross Earning",     ["gross", "gross earning", "gross earnings", "total earnings", "gross salary", "gross pay", "total gross"]),
    ("gross_deduction", "Gross Deduction",   ["totded", "tot ded", " totded", "total deduction", "total deductions", "gross deduction", "deductions"]),
    ("net_amount",      "Net Amount",        ["netsal", "net sal", "net salary", "net amount", "take home", "net pay", "in hand", "net payable", "net payout"]),
]

# ─── Multi-Company Registry ────────────────────────────────────────────────────
# Add / edit companies here. Keys are used as internal IDs in the UI.
COMPANIES = {
    "yayath": {
        "name": "YAYATH BUILDTECH",
        "address": "UG-006, MGF Metropolis Mall, MG Road, Sector 28 Gurugram, Haryana 122009",
        "logo": "logo.png",           # file inside static/img/
    },
    "adidev": {
        "name": "ADIDEV WAREHOUSE",
        "address": "UG-006, MGF Metropolis Mall, MG Road, Sector 28 Gurugram, Haryana 122009",
        "logo": "logo_adidev.png",
    },
    "shrikrishna": {
        "name": "SHRI KRISHNA WAREHOUSE",
        "address": "UG-006, MGF Metropolis Mall, MG Road, Sector 28 Gurugram, Haryana 122009",
        "logo": "logo_shrikrishna.png",
    },
}
DEFAULT_COMPANY_KEY  = "yayath"
DEFAULT_COMPANY = COMPANIES[DEFAULT_COMPANY_KEY]["name"]
DEFAULT_ADDRESS = COMPANIES[DEFAULT_COMPANY_KEY]["address"]

def norm(x):
    return re.sub(r"[^a-z0-9]+", " ", str(x).strip().lower()).strip()

def clean(v):
    if pd.isna(v):
        return ""
    if isinstance(v, (float, int)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()

def num(v):
    try:
        s = str(v).replace(",", "").replace("₹", "").replace("INR", "").strip()
        return float(s) if s and s.lower() not in ("nan", "none", "n/a", "-", "nil") else 0.0
    except Exception:
        return 0.0

def money(v):
    val = round(num(v))
    return f"{val:,.2f}"

def amount_words(value):
    n = int(round(num(value)))
    if n <= 0:
        return "ZERO ONLY"
    try:
        w = num2words(n, lang="en_IN").replace(",", "").upper()
        w = re.sub(r"\bAND\b", "", w)
        w = w.replace("-", " ")
        w = re.sub(r"\s+", " ", w).strip()
        return f"{w} ONLY"
    except Exception:
        try:
            w = num2words(n).replace(",", "").upper()
            w = re.sub(r"\bAND\b", "", w)
            w = w.replace("-", " ")
            w = re.sub(r"\s+", " ", w).strip()
            return f"{w} RUPEES ONLY"
        except Exception:
            return f"{n:,} RUPEES ONLY"

def auto_mapping(columns):
    mapping = {}
    normalized = {norm(c): c for c in columns}
    mapped_cols = set()

    # Pass 1: Exact matches against candidate aliases
    for key, label, aliases in FIELDS:
        candidates = [norm(label)] + [norm(a) for a in aliases]
        for candidate in candidates:
            if candidate in normalized and normalized[candidate] not in mapped_cols:
                mapping[key] = normalized[candidate]
                mapped_cols.add(normalized[candidate])
                break

    # Pass 2: Fuzzy / substring matches for remaining unmapped fields
    for key, label, aliases in FIELDS:
        if key in mapping:
            continue
        candidates = sorted([norm(label)] + [norm(a) for a in aliases], key=len, reverse=True)
        for candidate in candidates:
            if len(candidate) < 3:
                continue
            matched = False
            for norm_c, orig_c in normalized.items():
                if orig_c in mapped_cols:
                    continue
                if key == "name" and any(w in norm_c for w in ["bank", "company", "father", "husband"]):
                    continue
                if "net" in key and "gross" in norm_c:
                    continue
                if "gross" in key and "net" in norm_c:
                    continue
                # Do not map employee deductions to employer contributions
                if any(w in norm_c for w in ["empr", "employer", "emr"]) and not any(w in key for w in ["empr", "employer", "emr"]):
                    continue
                if candidate in norm_c or norm_c in candidate:
                    mapping[key] = orig_c
                    mapped_cols.add(orig_c)
                    matched = True
                    break
            if matched:
                break

    return mapping

def getval(row, mapping, key, default=""):
    col = mapping.get(key)
    if col and col in row.index:
        val = clean(row[col])
        return val if val != "" else default
    return default

def calculate_row_financials(row, mapping):
    """Calculates all earnings, employee deductions, and net take-home salary."""
    b_ent  = round(num(getval(row, mapping, "basic_entitled")))
    b_earn = round(num(getval(row, mapping, "basic_earned"))) or b_ent
    b_arr  = round(num(getval(row, mapping, "basic_arrears")))

    h_ent  = round(num(getval(row, mapping, "hra_entitled")))
    h_earn = round(num(getval(row, mapping, "hra_earned"))) or h_ent
    h_arr  = round(num(getval(row, mapping, "hra_arrears")))

    add_ent  = round(num(getval(row, mapping, "add_allowance")))
    add_earn = add_ent
    add_arr  = 0.0

    bon_ent  = round(num(getval(row, mapping, "bonus_entitled")))
    bon_earn = round(num(getval(row, mapping, "bonus_earned"))) or bon_ent
    bon_arr  = round(num(getval(row, mapping, "bonus_arrears")))

    # Deductions: Employee deductions only!
    pf_ded   = round(num(getval(row, mapping, "pf")))
    esi_ded  = round(num(getval(row, mapping, "esi")))
    pt_ded   = round(num(getval(row, mapping, "prof_tax")))

    # Gross Earnings: Sum of rounded components so table always adds up perfectly
    gross_ent  = b_ent + h_ent + add_ent + bon_ent
    gross_earn = b_earn + b_arr + h_earn + h_arr + add_earn + add_arr + bon_earn + bon_arr

    # Gross Deductions: Sum of employee deductions
    gross_ded = pf_ded + esi_ded + pt_ded

    # Net Salary: Gross - Deduction
    net = gross_earn - gross_ded

    return {
        "basic_ent": b_ent, "basic_earn": b_earn, "basic_arr": b_arr,
        "hra_ent": h_ent, "hra_earn": h_earn, "hra_arr": h_arr,
        "add_ent": add_ent, "add_earn": add_earn, "add_arr": add_arr,
        "bonus_ent": bon_ent, "bonus_earn": bon_earn, "bonus_arr": bon_arr,
        "pf_ded": pf_ded,
        "esi_ded": esi_ded,
        "pt_ded": pt_ded,
        "gross_ent": gross_ent,
        "gross_earn": gross_earn,
        "gross_ded": gross_ded,
        "net": net
    }

def build_pdf(row, mapping, month, company=DEFAULT_COMPANY, address=DEFAULT_ADDRESS):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm
    )
    styles = getSampleStyleSheet()

    # Premium Typography Styles
    title_style = ParagraphStyle(
        "CompanyTitle",
        parent=styles["Heading1"],
        alignment=TA_CENTER,
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#0F172A"),
        fontName="Helvetica-Bold",
        spaceAfter=2
    )
    address_style = ParagraphStyle(
        "CompanyAddress",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#475569"),
        fontName="Helvetica",
        spaceAfter=6
    )
    month_badge_style = ParagraphStyle(
        "MonthBadge",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=10,
        leading=13,
        textColor=colors.HexColor("#1E3A8A"),
        fontName="Helvetica-Bold",
        spaceBefore=4,
        spaceAfter=10
    )
    normal_text = ParagraphStyle(
        "NormalText",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=10.5,
        textColor=colors.HexColor("#1E293B"),
        fontName="Helvetica"
    )
    table_header_style = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=10,
        textColor=colors.HexColor("#0F172A"),
        fontName="Helvetica-Bold"
    )
    table_right_header = ParagraphStyle(
        "TableRightHeader",
        parent=table_header_style,
        alignment=TA_RIGHT
    )
    cell_bold = ParagraphStyle(
        "CellBold",
        parent=normal_text,
        fontName="Helvetica-Bold"
    )
    cell_right = ParagraphStyle(
        "CellRight",
        parent=normal_text,
        alignment=TA_RIGHT
    )
    cell_bold_right = ParagraphStyle(
        "CellBoldRight",
        parent=normal_text,
        alignment=TA_RIGHT,
        fontName="Helvetica-Bold"
    )
    net_label_style = ParagraphStyle(
        "NetLabel",
        parent=styles["Normal"],
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#0F172A"),
        fontName="Helvetica-Bold"
    )
    net_val_style = ParagraphStyle(
        "NetValue",
        parent=styles["Normal"],
        alignment=TA_RIGHT,
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#0F172A"),
        fontName="Helvetica-Bold"
    )
    words_style = ParagraphStyle(
        "WordsStyle",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=10.5,
        textColor=colors.HexColor("#1E293B"),
        fontName="Helvetica"
    )
    footer_note = ParagraphStyle(
        "FooterNote",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=7,
        leading=9.5,
        textColor=colors.HexColor("#64748B"),
        fontName="Helvetica-Bold"
    )

    story = []

    # 1. Company Header — only show logo for YAYATH BUILDTECH
    YAYATH_KEY = "yayath"
    is_yayath = company.strip().upper() == COMPANIES[YAYATH_KEY]["name"].upper()

    if is_yayath:
        logo_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "static", "img", "logo.png"
        )
        if os.path.exists(logo_path):
            try:
                logo_img = RLImage(logo_path, width=38 * mm, height=17.5 * mm)
                logo_img.hAlign = 'CENTER'
                story.append(logo_img)
                story.append(Spacer(1, 2 * mm))
            except Exception:
                pass

    story.append(Paragraph(company, title_style))
    story.append(Paragraph(address, address_style))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#E5A93B"), spaceBefore=3, spaceAfter=6))
    story.append(Paragraph(f"Payslip for the month of {month}", month_badge_style))

    # 2. Employee Info & Attendance Grid
    emp_code = getval(row, mapping, "code") or row.get("sno", "") or ""
    paid_d   = getval(row, mapping, "paid_days") or "30"
    lwp_d    = getval(row, mapping, "lwp") or "0"

    pres_d = getval(row, mapping, "present")
    if not pres_d:
        try:
            pd_float = float(paid_d)
            pres_d = f"{(pd_float - 4.0):.2f}" if pd_float >= 28 else f"{pd_float:.2f}"
        except Exception:
            pres_d = "27.00"
    else:
        try:
            pres_d = f"{float(pres_d):.2f}"
        except Exception:
            pass

    left_fields = [
        ("Name",          getval(row, mapping, "name")),
        ("Employee Code", emp_code),
        ("Department",    getval(row, mapping, "department") or "OPERATION"),
        ("Designation",   getval(row, mapping, "designation")),
        ("PAN",           getval(row, mapping, "pan")),
        ("Leave Bal.",    getval(row, mapping, "leave_bal")),
        ("ESI No.",       getval(row, mapping, "esi_no", "N/A")),
    ]
    right_fields = [
        ("Location",         getval(row, mapping, "location") or "BARODA"),
        ("Unit",             getval(row, mapping, "unit") or "BARODA"),
        ("Joining Date",     getval(row, mapping, "joining_date") or ""),
        ("Paid days",        f"{float(paid_d):.2f}" if paid_d and str(paid_d).replace('.', '', 1).isdigit() else paid_d),
        ("Bank Account No.", getval(row, mapping, "bank", "N/A")),
        ("PF No.",           getval(row, mapping, "pf_no")),
        ("Cost Center",      getval(row, mapping, "cost_center", company)),
    ]
    attend_fields = [
        ("Present", pres_d),
        ("LWP",     f"{float(lwp_d):.2f}" if lwp_d and str(lwp_d).replace('.', '', 1).isdigit() else "0.00"),
        ("NH",      f"{float(getval(row, mapping, 'nh', '0')):.2f}" if getval(row, mapping, 'nh') and str(getval(row, mapping, 'nh')).replace('.', '', 1).isdigit() else "0.00"),
        ("WOFF",    f"{float(getval(row, mapping, 'woff', '4')):.2f}" if getval(row, mapping, 'woff') and str(getval(row, mapping, 'woff')).replace('.', '', 1).isdigit() else "4.00"),
        ("", ""),
        ("", ""),
        ("", "")
    ]

    max_rows = max(len(left_fields), len(right_fields), len(attend_fields))
    detail_data = []
    for i in range(max_rows):
        l = left_fields[i] if i < len(left_fields) and left_fields[i][0] else ("", "")
        r = right_fields[i] if i < len(right_fields) and right_fields[i][0] else ("", "")
        att = attend_fields[i] if i < len(attend_fields) and attend_fields[i][0] else ("", "")

        col1_text = f"<b>{l[0]}</b> :{l[1]}" if l[0] else ""
        col2_text = f"<b>{r[0]}</b> :{r[1]}" if r[0] else ""
        col3_text = f"<b>{att[0]}</b> :{att[1]}" if att[0] else ""

        detail_data.append([
            Paragraph(col1_text, normal_text),
            Paragraph(col2_text, normal_text),
            Paragraph(col3_text, normal_text)
        ])

    detail_table = Table(detail_data, colWidths=[68 * mm, 72 * mm, 44 * mm])
    detail_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E2E8F0")),
    ]))
    story.append(detail_table)
    story.append(Spacer(1, 8))

    # 3. Earnings & Deductions Calculations — Exact 6-column format
    fin = calculate_row_financials(row, mapping)

    earnings_table_data = [
        [
            Paragraph("<b>Earnings</b>", table_header_style),
            Paragraph("<b>Entitled Amt.</b>", table_right_header),
            Paragraph("<b>Earned Amt.</b>", table_right_header),
            Paragraph("<b>Arrears</b>", table_right_header),
            Paragraph("<b>Deductions</b>", table_header_style),
            Paragraph("<b>Amount</b>", table_right_header),
        ],
        [
            Paragraph("BASIC", normal_text),
            Paragraph(money(fin["basic_ent"]), cell_right),
            Paragraph(money(fin["basic_earn"]), cell_right),
            Paragraph(money(fin["basic_arr"]), cell_right),
            Paragraph("PF", normal_text),
            Paragraph(money(fin["pf_ded"]), cell_right)
        ],
        [
            Paragraph("HRA", normal_text),
            Paragraph(money(fin["hra_ent"]), cell_right),
            Paragraph(money(fin["hra_earn"]), cell_right),
            Paragraph(money(fin["hra_arr"]), cell_right),
            Paragraph("ESI", normal_text),
            Paragraph(money(fin["esi_ded"]), cell_right)
        ],
    ]

    if fin.get("add_earn", 0) > 0:
        earnings_table_data.append([
            Paragraph("Additional Allowance", normal_text),
            Paragraph(money(fin["add_ent"]), cell_right),
            Paragraph(money(fin["add_earn"]), cell_right),
            Paragraph(money(fin["add_arr"]), cell_right),
            Paragraph("", normal_text),
            Paragraph("", cell_right)
        ])

    earnings_table_data.append([
        Paragraph("Bonus", normal_text),
        Paragraph(money(fin["bonus_ent"]), cell_right),
        Paragraph(money(fin["bonus_earn"]), cell_right),
        Paragraph(money(fin["bonus_arr"]), cell_right),
        Paragraph("Prof. Tax", normal_text),
        Paragraph(money(fin["pt_ded"]), cell_right)
    ])

    earnings_table_data.append([
        Paragraph("<b>Gross Earning</b>", cell_bold),
        Paragraph(f"<b>{money(fin['gross_ent'])}</b>", cell_bold_right),
        Paragraph(f"<b>{money(fin['gross_earn'])}</b>", cell_bold_right),
        Paragraph("", cell_right),
        Paragraph("<b>Gross Deduction</b>", cell_bold),
        Paragraph(f"<b>{money(fin['gross_ded'])}</b>", cell_bold_right)
    ])

    # Total width = 38 + 28 + 28 + 24 + 38 + 28 = 184mm
    et = Table(earnings_table_data, colWidths=[38 * mm, 28 * mm, 28 * mm, 24 * mm, 38 * mm, 28 * mm], repeatRows=1)
    et.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ("ALIGN", (1, 0), (3, -1), "RIGHT"),
        ("ALIGN", (5, 0), (5, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94A3B8")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F1F5F9")),
    ]))
    story.append(et)
    story.append(Spacer(1, 8))

    # 4. Net Amount Card
    net_val = fin["net"]
    net_table_data = [
        [
            Paragraph("<b>Net Amount</b>", net_label_style),
            Paragraph(f"<b>{money(net_val)} INR</b>", net_val_style)
        ]
    ]
    net_table = Table(net_table_data, colWidths=[120 * mm, 64 * mm])
    net_table.setStyle(TableStyle([
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FEF3C7")),
        ("BOX", (0, 0), (-1, -1), 1.0, colors.HexColor("#E5A93B")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(net_table)
    story.append(Spacer(1, 6))

    # 5. Amount in Words & Notes
    words = amount_words(net_val)
    story.append(Paragraph(f"<b>Net Amount in words:</b> ( {words} )", words_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("(*) Payhead which are not part of gross", ParagraphStyle("StarNote", parent=normal_text, fontSize=6.5, textColor=colors.HexColor("#64748B"))))
    story.append(Spacer(1, 10))

    # 6. Legal / System Notice & Perforated Cut line
    story.append(Paragraph("THIS IS A SYSTEM GENERATED PAYSLIP, DOES NOT REQUIRE ANY SIGNATURE AND/OR COMPANY SEAL.", footer_note))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#94A3B8"), spaceAfter=4))
    story.append(Paragraph("<para alignment='center'>------------------ Cut Here --------------------</para>", ParagraphStyle("Cut", parent=footer_note, fontSize=6.5)))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

@app.route("/login", methods=["GET", "POST"])
def login():
    # Already logged in? Go home
    if get_current_user():
        return redirect(url_for("index"))

    error = None
    if request.method == "POST":
        uid = request.form.get("user_id", "").strip()
        pwd = request.form.get("password", "").strip()

        if uid == AUTH_USER and pwd == AUTH_PASS:
            token = _make_jwt(uid)
            resp = make_response(redirect(url_for("index")))
            resp.set_cookie(
                JWT_COOKIE,
                token,
                httponly=True,
                samesite="Lax",
                secure=IS_VERCEL,
                max_age=JWT_EXPIRY_HRS * 3600,
                path="/"
            )
            return resp
        else:
            error = "Invalid User ID or Password. Please try again."

    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    resp = make_response(redirect(url_for("login")))
    resp.delete_cookie(JWT_COOKIE, path="/")
    return resp

@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or not f.filename:
            flash("Please choose an Excel or CSV file to upload.", "warning")
            return redirect(url_for("index"))

        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED:
            flash("Invalid format. Please upload .xlsx, .xls or .csv only.", "danger")
            return redirect(url_for("index"))

        try:
            if ext == ".csv":
                df = pd.read_csv(f, dtype=str).fillna("")
            else:
                # Read without assuming row 0 is header to handle top banner/title rows
                raw_df = pd.read_excel(f, dtype=str, header=None).fillna("")
                header_row_idx = 0
                max_score = -1
                for r_idx in range(min(8, len(raw_df))):
                    row_vals = [norm(str(v)) for v in raw_df.iloc[r_idx] if str(v).strip()]
                    if not row_vals:
                        continue
                    # Count how many payroll keywords match this row
                    score = sum(1 for v in row_vals if any(k in v for k in [
                        "name", "code", "emp", "basic", "salary", "gross", "net",
                        "ded", "pf", "esi", "days", "pan", "dept", "desig", "leave"
                    ]))
                    if score > max_score:
                        max_score = score
                        header_row_idx = r_idx

                if max_score >= 2 and header_row_idx > 0:
                    df = raw_df.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
                    df.columns = [str(c).strip() for c in raw_df.iloc[header_row_idx]]
                    # Drop unnamed or empty columns
                    df = df.loc[:, [bool(c) and not str(c).startswith("Unnamed") for c in df.columns]]
                else:
                    f.seek(0)
                    df = pd.read_excel(f, dtype=str).fillna("")
                df.columns = [str(c).strip() for c in df.columns]

            # Filter out blank rows or trailing Total/Summary rows
            if not df.empty:
                df = df.dropna(how="all")
                def is_total_row(row):
                    text = " ".join(str(v).lower() for v in row.values)
                    return any(k in text for k in ["total", "grand total", "sum"])
                df = df[~df.apply(is_total_row, axis=1)]
                cols_to_check = [c for c in df.columns if any(k in c.lower() for k in ["name", "desig", "basic", "emp"])]
                if cols_to_check:
                    df = df[df[cols_to_check].apply(lambda r: "".join(str(x).strip() for x in r) != "", axis=1)]
                df = df.reset_index(drop=True)
        except Exception as e:
            flash(f"Could not read spreadsheet file: {e}", "danger")
            return redirect(url_for("index"))

        if df.empty:
            flash("The uploaded sheet has no records or data rows.", "warning")
            return redirect(url_for("index"))

        columns = list(df.columns)
        mapping = auto_mapping(columns)

        # Fallback for empty employee codes using sno or row index
        code_col = mapping.get("code")
        sno_col = next((c for c in columns if c.lower() in ["sno", "s.no", "sr no", "sr.no"]), None)
        if code_col and code_col in df.columns:
            for idx in range(len(df)):
                val = str(df.at[idx, code_col]).strip()
                if not val or val.lower() == "nan":
                    if sno_col and str(df.at[idx, sno_col]).strip():
                        df.at[idx, code_col] = str(df.at[idx, sno_col]).strip()
                    else:
                        df.at[idx, code_col] = f"{idx+1:03d}"

        records = df.to_dict(orient="records")

        # Compute summary metrics for preview banner
        total_employees = len(records)
        month = request.form.get("month") or datetime.now().strftime("%B, %Y")
        company = request.form.get("company") or DEFAULT_COMPANY
        address = request.form.get("address") or DEFAULT_ADDRESS

        # Calculate quick summary aggregates
        total_payout = 0.0
        total_gross = 0.0
        total_ded = 0.0
        for r in records:
            fin = calculate_row_financials(pd.Series(r), mapping)
            total_payout += fin["net"]
            total_gross += fin["gross_earn"]
            total_ded += fin["gross_ded"]

        summary = {
            "total_employees": total_employees,
            "total_payout": f"{total_payout:,.2f}",
            "total_gross": f"{total_gross:,.2f}",
            "total_ded": f"{total_ded:,.2f}",
            "mapped_count": len([k for k, v in mapping.items() if v in columns]),
            "total_fields": len(FIELDS)
        }

        # Build a single JSON blob for safe embedding in a <script> tag
        # (avoids Jinja2 HTML auto-escaping of quotes in hidden input values)
        payroll_data_json = json.dumps({
            "records": records,
            "mapping": mapping,
            "month": month,
            "company": company,
            "address": address
        }, ensure_ascii=False)

        return render_template(
            "mapping.html",
            columns=columns,
            fields=FIELDS,
            mapping=mapping,
            records=records,
            summary=summary,
            month=month,
            company=company,
            address=address,
            payroll_data_json=payroll_data_json
        )

    # GET: render upload dashboard
    current_month_str = datetime.now().strftime("%B, %Y")
    return render_template(
        "index.html",
        companies=COMPANIES,
        default_company_key=DEFAULT_COMPANY_KEY,
        default_company=DEFAULT_COMPANY,
        default_address=DEFAULT_ADDRESS,
        current_month=current_month_str
    )

@app.route("/preview-pdf", methods=["POST"])
@login_required
def preview_pdf():
    """Generates a single employee PDF and returns it inline for modal iframe viewing."""
    try:
        data = request.get_json(force=True)
        row_dict = data.get("record", {})
        mapping = data.get("mapping", {})
        month = data.get("month", "") or datetime.now().strftime("%B, %Y")
        company = data.get("company", DEFAULT_COMPANY)
        address = data.get("address", DEFAULT_ADDRESS)

        row = pd.Series(row_dict)
        pdf_bytes = build_pdf(row, mapping, month, company, address)

        name = getval(row, mapping, "name") or "Employee"
        code = getval(row, mapping, "code") or "001"
        safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{name}_{code}").strip("_")

        return Response(
            pdf_bytes,
            mimetype="application/pdf",
            headers={"Content-Disposition": f"inline; filename={safe_name}.pdf"}
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    try:
        records = json.loads(request.form["records"])
        mapping = json.loads(request.form["mapping"])
    except Exception:
        flash("Could not process the uploaded data. Please upload the file again.", "danger")
        return redirect(url_for("index"))

    month = request.form.get("month", "").strip() or datetime.now().strftime("%B, %Y")
    company = request.form.get("company", DEFAULT_COMPANY).strip() or DEFAULT_COMPANY
    address = request.form.get("address", DEFAULT_ADDRESS).strip() or DEFAULT_ADDRESS
    selected = request.form.get("employee_index", "all")

    rows = list(enumerate(records))
    if selected != "all":
        try:
            sel_idx = int(selected)
            rows = [(sel_idx, records[sel_idx])]
        except Exception:
            rows = []

    if not rows:
        flash("No employee selected for payslip generation.", "warning")
        return redirect(url_for("index"))

    outputs = []
    for idx, rowdict in rows:
        row = pd.Series(rowdict)
        code = getval(row, mapping, "code") or str(idx + 1)
        name = getval(row, mapping, "name") or f"Employee_{idx + 1}"
        safe = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{name}_{code}").strip("_")
        pdf = build_pdf(row, mapping, month, company, address)
        outputs.append((f"{safe}.pdf", pdf))

    # Single employee download
    if len(outputs) == 1:
        return send_file(
            io.BytesIO(outputs[0][1]),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=outputs[0][0]
        )

    # Bulk ZIP download
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in outputs:
            archive.writestr(name, data)
    z.seek(0)

    clean_month = re.sub(r"[^A-Za-z0-9]+", "_", month)
    zip_name = f"Yayath_Salary_Slips_{clean_month}.zip"
    return send_file(
        z,
        mimetype="application/zip",
        as_attachment=True,
        download_name=zip_name
    )

@app.route("/download-sample", methods=["GET"])
@login_required
def download_sample():
    """Generates and serves a realistic pre-filled sample Excel spreadsheet for instant testing."""
    sample_data = [
        {
            "Employee Code": "YBT-101",
            "Employee Name": "Rajesh Sharma",
            "Department": "Civil Engineering",
            "Designation": "Senior Project Manager",
            "PAN": "ABCDE1234F",
            "Joining Date": "2021-04-15",
            "Location": "Gurugram",
            "Unit": "Unit-I",
            "Leave Balance": "5",
            "Present Days": "26",
            "LWP": "0",
            "NH": "2",
            "WOFF": "4",
            "Paid Days": "30",
            "Bank Account No.": "98765432101234",
            "PF No.": "100234567891",
            "ESI No.": "N/A",
            "Cost Center": "YAYATH BUILDTECH",
            "Basic Entitled": "45000",
            "Basic Earned": "45000",
            "Basic Arrears": "0",
            "HRA Entitled": "18000",
            "HRA Earned": "18000",
            "HRA Arrears": "0",
            "Bonus Entitled": "5000",
            "Bonus Earned": "5000",
            "Bonus Arrears": "0",
            "PF Deduction": "1800",
            "ESI Deduction": "0",
            "Professional Tax": "200",
            "Gross Earning": "68000",
            "Gross Deduction": "2000",
            "Net Amount": "66000"
        },
        {
            "Employee Code": "YBT-102",
            "Employee Name": "Pooja Verma",
            "Department": "Architecture & Design",
            "Designation": "Lead Architect",
            "PAN": "BCDEF2345G",
            "Joining Date": "2022-08-01",
            "Location": "Gurugram",
            "Unit": "Unit-I",
            "Leave Balance": "4",
            "Present Days": "25",
            "LWP": "1",
            "NH": "2",
            "WOFF": "4",
            "Paid Days": "29",
            "Bank Account No.": "54321098765432",
            "PF No.": "100345678912",
            "ESI No.": "N/A",
            "Cost Center": "YAYATH BUILDTECH",
            "Basic Entitled": "40000",
            "Basic Earned": "38667",
            "Basic Arrears": "0",
            "HRA Entitled": "16000",
            "HRA Earned": "15467",
            "HRA Arrears": "0",
            "Bonus Entitled": "4000",
            "Bonus Earned": "3867",
            "Bonus Arrears": "0",
            "PF Deduction": "1800",
            "ESI Deduction": "0",
            "Professional Tax": "200",
            "Gross Earning": "58001",
            "Gross Deduction": "2000",
            "Net Amount": "56001"
        },
        {
            "Employee Code": "YBT-103",
            "Employee Name": "Amit Kumar Patel",
            "Department": "Site Operations",
            "Designation": "Site Supervisor",
            "PAN": "CDEFG3456H",
            "Joining Date": "2023-01-10",
            "Location": "Noida Site",
            "Unit": "Unit-II",
            "Leave Balance": "2",
            "Present Days": "26",
            "LWP": "0",
            "NH": "2",
            "WOFF": "4",
            "Paid Days": "30",
            "Bank Account No.": "65432109876543",
            "PF No.": "100456789123",
            "ESI No.": "210987654321",
            "Cost Center": "YAYATH BUILDTECH",
            "Basic Entitled": "22000",
            "Basic Earned": "22000",
            "Basic Arrears": "1500",
            "HRA Entitled": "8800",
            "HRA Earned": "8800",
            "HRA Arrears": "0",
            "Bonus Entitled": "2500",
            "Bonus Earned": "2500",
            "Bonus Arrears": "0",
            "PF Deduction": "1800",
            "ESI Deduction": "246",
            "Professional Tax": "200",
            "Gross Earning": "34800",
            "Gross Deduction": "2246",
            "Net Amount": "32554"
        },
        {
            "Employee Code": "YBT-104",
            "Employee Name": "Neha Gupta",
            "Department": "Finance & Accounts",
            "Designation": "Senior Accountant",
            "PAN": "DEFGH4567I",
            "Joining Date": "2022-11-15",
            "Location": "Gurugram Head Office",
            "Unit": "Unit-I",
            "Leave Balance": "6",
            "Present Days": "26",
            "LWP": "0",
            "NH": "2",
            "WOFF": "4",
            "Paid Days": "30",
            "Bank Account No.": "76543210987654",
            "PF No.": "100567891234",
            "ESI No.": "N/A",
            "Cost Center": "YAYATH BUILDTECH",
            "Basic Entitled": "35000",
            "Basic Earned": "35000",
            "Basic Arrears": "0",
            "HRA Entitled": "14000",
            "HRA Earned": "14000",
            "HRA Arrears": "0",
            "Bonus Entitled": "3500",
            "Bonus Earned": "3500",
            "Bonus Arrears": "0",
            "PF Deduction": "1800",
            "ESI Deduction": "0",
            "Professional Tax": "200",
            "Gross Earning": "52500",
            "Gross Deduction": "2000",
            "Net Amount": "50500"
        }
    ]

    df_sample = pd.DataFrame(sample_data)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_sample.to_excel(writer, index=False, sheet_name="Payroll Data")
    out.seek(0)

    return send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Yayath_Payroll_Sample_Template.xlsx"
    )

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
