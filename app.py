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
    ("name",            "Employee Name",         ["empname", "emp name", "employee name", "employee_name", "name", "full name", "staff name", "worker name"]),
    ("code",            "Employee Code",         ["empcode", "emp code", "employee code", "employee_code", "emp_code", "code", "employee id", "emp id", "id", "staff no", "badge no"]),
    ("designation",     "Designation",           ["designation", "role", "position", "title", "post", "job title"]),
    ("department",      "Department",            ["department", "dept", "department name", "dept name", "section"]),
    ("pan",             "PAN",                   ["pan", "pan number", "pan_no", "pan card", "panno", "pan no."]),
    ("joining_date",    "Joining Date",          ["joining date", "joining_date", "date of joining", "doj", "date_of_joining"]),
    ("location",        "Location / Site",       ["location", "city", "work location", "site", "project site", "branch location", "project"]),
    ("unit",            "Unit",                  ["unit", "branch", "business unit", "division"]),
    ("cost_center",     "Cost Center",           ["cost center", "cost_center", "cc"]),

    # ── Employee IDs ─────────────────────────────────────────────────────────
    ("bank",            "Bank Account No.",      ["bank account no", "bank account", "bank_account", "account number", "a/c no", "account_no", "bank a/c", "acc no"]),
    ("pf_no",           "PF No. / UAN",          ["pf_no", "pf no", "pf number", "uan", "pf uan", "provident fund no", "uan no", "pf a/c"]),
    ("esi_no",          "ESI No.",               ["esi no", "esi no.", "esi number", "esi_no", "insurance no", "esic no", "ip number", "esic"]),
    ("leave_bal",       "Leave Balance",         ["leave bal", "leave balance", "leave_bal", "leaves", "balance leaves", "cl", "pl", "bal leaves"]),

    # ── Attendance ────────────────────────────────────────────────────────────
    ("total_days",      "Total Days",            ["totaldays", "total days", "total_days", "working days in month", "calendar days", "days"]),
    ("paid_days",       "Paid Days",             ["paiddays", "paid days", "paid_days", "payable days", "net paid days"]),
    ("lwp",             "Leave Without Pay",     ["leave without pay", "lwp", "lop", "loss of pay", "unpaid", "absent", "leave_without_pay"]),
    ("present",         "Present Days",          ["present", "present days", "days worked", "total present"]),
    ("nh",              "NH",                    ["nh", "national holiday", "holidays", "festival holiday"]),
    ("woff",            "WOFF",                  ["woff", "weekly off", "weekly_off", "off days", "wo", "week off"]),
    ("extra_day_att",   "Extra Days (Att)",      ["exra day", "extra day", "extra days", "2 oct - extra pay", "extra pay day", "ed days", "holiday day", "extra working day", "exra days"]),
    ("ot_hours",        "OT Hours / Days",       ["ot hours", "ot hrs", "ot days", "ot day", "overtime hours", "overtime days"]),

    # ── Earnings ─────────────────────────────────────────────────────────────
    ("basic_entitled",  "Basic Entitled",        ["basic da", "basic +da", "basic+da", "basic", "basic salary", "entitled basic", "basic pay", "basic rate", "basic entitled", "basic + da"]),
    ("basic_earned",    "Basic Earned",          ["basic earned", "earned basic", "basic pay earned"]),
    ("basic_arrears",   "Basic Arrears",         ["basic arrears", "arrears basic", "basic arrear"]),
    ("hra_entitled",    "HRA Entitled",          ["hra", "house rent allowance", "hra rate", "hra entitled"]),
    ("hra_earned",      "HRA Earned",            ["hra earned", "earned hra"]),
    ("hra_arrears",     "HRA Arrears",           ["hra arrears", "arrears hra"]),
    ("add_allowance",   "Additional Allowance",  ["additional allowance", "additional_allowance", "add allowance", "addl allowance", "special allowance", "other allowance"]),
    ("conveyance",      "Conveyance Allowance",  ["local conveyance", "conveyance", "conveyance allowance", "transport allowance", "travel allowance", "ta"]),
    ("washing",         "Washing Allowance",     ["washing allowance", "washing", "uniform allowance", "dress allowance"]),
    ("leave_encashment","Leave Encashment",      ["attendance leave encashment", "leave encashment", "leave encash", "encashment", "el encashment", "pl encashment", "leave pay"]),
    ("incentive",       "Incentive / Festival",  ["navratri festival incentive", "festival incentive", "incentive", "performance incentive", "diwali incentive", "special incentive", "attendance incentive", "festival bonus"]),
    ("extra_day_earn",  "Extra Day Pay",         ["exra day .1", "exra day.1", "extra day .1", "extra day.1", "2 oct - extra pay .1", "2 oct - extra pay.1", "extra day amount", "extra day pay", "extra days pay", "ed amount", "extra pay"]),
    ("ot",              "OT Amount",             ["ot", "ot amount", "overtime", "overtime amount", "ot pay", "over time", "overtime pay", "ot wages"]),
    ("bonus_entitled",  "Bonus Entitled",        ["bonus", "statutory bonus", "bonus entitled", "bonus rate"]),
    ("bonus_earned",    "Bonus Earned",          ["bonus earned", "earned bonus"]),
    ("bonus_arrears",   "Bonus Arrears",         ["bonus arrears", "arrears bonus"]),

    # ── Deductions ───────────────────────────────────────────────────────────
    ("pf",              "PF Deduction",          ["pf", "pf ", "pf deduction", "pf amount", "provident fund", "epf", "pf ee", "pf employee"]),
    ("esi",             "ESI Deduction",         ["esi basic da", "esi basic", "esi", "esi deduction", "esi amount", "esic ee", "esi employee", "esi (basic +da)", "esi (basic+da)"]),
    ("prof_tax",        "Professional Tax",      ["proftax", "prof tax", "professional tax", "pt", "p tax", "ptax"]),
    ("lwf",             "Labour Welfare Fund",   ["lwf- em", "lwf - em", "lwf-em", "lwf employee", "lwf ded", "labour welfare fund", "lwf", "lwf em"]),

    # ── Totals (optional) ────────────────────────────────────────────────────
    ("gross_earning",   "Gross Earning",         ["grosstotal", "gross total", "gross esi", "gross_esi", "gross earning", "gross earnings", "total earnings", "gross salary", "gross pay", "gross"]),
    ("gross_deduction", "Gross Deduction",       ["totded", "tot ded", " totded", "total deduction", "total deductions", "gross deduction", "deductions"]),
    ("net_amount",      "Net Amount",            ["netsal", "net sal", "net salary", "net amount", "take home", "net pay", "in hand", "net payable"]),
]

# ─── Multi-Company Registry ────────────────────────────────────────────────────
# Add / edit companies here. Keys are used as internal IDs in the UI.
COMPANIES = {
    "yayath": {
        "name": "YAYATH BUILDTECH",
        "address": "UG-006, MGF Metropolis Mall, MG Road, Sector 28 Gurugram, Haryana 122009",
        "logo": "logo.png",
    },
    "adidev": {
        "name": "ADIDEV WAREHOUSE",
        "address": "UG-006, MGF Metropolis Mall, MG Road, Sector 28 Gurugram, Haryana 122009",
        "logo": "logo.png",
    },
    "shrikrishna": {
        "name": "SHRI KRISHNA WAREHOUSE",
        "address": "UG-006, MGF Metropolis Mall, MG Road, Sector 28 Gurugram, Haryana 122009",
        "logo": "logo.png",
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

def detect_location(banner_text: str = "", filename: str = "") -> str:
    """Auto-detect site / location from uploaded file banner or filename."""
    CITIES = [
        "GOA", "RAJKOT", "BARODA", "VADODARA", "SURAT", "AHMEDABAD",
        "GURUGRAM", "GURGAON", "NOIDA", "MUMBAI", "PUNE", "DELHI",
        "JAIPUR", "INDORE", "BENGALURU", "BANGALORE", "HYDERABAD", "CHENNAI", "KOLKATA"
    ]
    # 1. From banner text (e.g. "Adidev - Goa")
    if banner_text:
        bt_upper = str(banner_text).upper()
        if "-" in bt_upper:
            parts = [p.strip() for p in bt_upper.split("-") if p.strip()]
            if len(parts) > 1:
                clean_cand = re.sub(r"\b(WAGE|SHEET|WAGESS|OCT|NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|\d+)\b", "", parts[-1]).strip()
                if clean_cand:
                    return "BARODA" if clean_cand == "VADODARA" else clean_cand
        for city in CITIES:
            if re.search(rf"\b{city}\b", bt_upper):
                return "BARODA" if city == "VADODARA" else city

    # 2. From filename (e.g. "GOA_-_WAGE_SHEET_-_Oct_26.xlsx")
    fn_upper = os.path.basename(str(filename)).upper()
    for city in CITIES:
        if re.search(rf"\b{city}\b", fn_upper) or city in fn_upper:
            return "BARODA" if city == "VADODARA" else city

    # 3. Clean banner fallback
    if banner_text:
        clean_b = re.sub(r"[^A-Za-z0-9\s-]", " ", str(banner_text)).strip()
        if clean_b:
            last = clean_b.split()[-1].upper()
            if len(last) >= 3 and not any(k in last for k in ["WAGE", "SHEET", "DATA"]):
                return last

    return ""

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
                norm_c = candidate
                if any(w in norm_c for w in ["empr", "employer", "emr"]) and not any(w in key for w in ["empr", "employer", "emr"]):
                    continue
                if "ctc" in norm_c and not "ctc" in key:
                    continue
                orig_c = normalized[candidate]
                mapping[key] = orig_c
                mapped_cols.add(orig_c)
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
                if "net" in key and ("gross" in norm_c or "ctc" in norm_c):
                    continue
                if "gross" in key and ("net" in norm_c or "ctc" in norm_c):
                    continue
                if any(w in norm_c for w in ["empr", "employer", "emr"]) and not any(w in key for w in ["empr", "employer", "emr"]):
                    continue
                if "ctc" in norm_c and not "ctc" in key:
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
    """Calculates all earnings, employee deductions, and net take-home salary.
    Ensures Net Salary ALWAYS accurately matches NetSal from the spreadsheet."""
    b_ent  = round(num(getval(row, mapping, "basic_entitled", 0)))
    b_earn = round(num(getval(row, mapping, "basic_earned", 0))) or b_ent
    b_arr  = round(num(getval(row, mapping, "basic_arrears", 0)))

    h_ent  = round(num(getval(row, mapping, "hra_entitled", 0)))
    h_earn = round(num(getval(row, mapping, "hra_earned", 0))) or h_ent
    h_arr  = round(num(getval(row, mapping, "hra_arrears", 0)))

    add_ent  = round(num(getval(row, mapping, "add_allowance", 0)))
    add_earn = add_ent
    add_arr  = 0.0

    conv_earn      = round(num(getval(row, mapping, "conveyance", 0)))
    wash_earn      = round(num(getval(row, mapping, "washing", 0)))
    leave_enc_earn = round(num(getval(row, mapping, "leave_encashment", 0)))
    incentive_earn = round(num(getval(row, mapping, "incentive", 0)))
    extra_day_earn = round(num(getval(row, mapping, "extra_day_earn", getval(row, mapping, "extra_day", 0))))
    ot_earn        = round(num(getval(row, mapping, "ot", 0)))

    bon_ent  = round(num(getval(row, mapping, "bonus_entitled", 0)))
    bon_earn = round(num(getval(row, mapping, "bonus_earned", 0))) or bon_ent
    bon_arr  = round(num(getval(row, mapping, "bonus_arrears", 0)))

    # Deductions: Employee deductions only!
    pf_ded   = round(num(getval(row, mapping, "pf", 0)))
    esi_ded  = round(num(getval(row, mapping, "esi", 0)))
    pt_ded   = round(num(getval(row, mapping, "prof_tax", 0)))
    lwf_ded  = round(num(getval(row, mapping, "lwf", 0)))

    calc_ded = pf_ded + esi_ded + pt_ded + lwf_ded

    # Check if Gross Deduction is explicitly mapped and provided in sheet
    sheet_ded_raw = getval(row, mapping, "gross_deduction", "")
    if sheet_ded_raw != "" and num(sheet_ded_raw) > 0:
        gross_ded = round(num(sheet_ded_raw))
    else:
        gross_ded = calc_ded

    # Sum of individual earnings components
    calc_gross_earn = b_earn + b_arr + h_earn + h_arr + add_earn + add_arr + conv_earn + wash_earn + leave_enc_earn + incentive_earn + extra_day_earn + ot_earn + bon_earn + bon_arr

    # Check if NetSal / net_amount is explicitly mapped and provided in sheet
    sheet_net_raw = getval(row, mapping, "net_amount", "")
    sheet_gross_raw = getval(row, mapping, "gross_earning", "")

    if sheet_net_raw != "" and num(sheet_net_raw) > 0:
        # User's sheet has explicit NetSal — match NetSal exactly!
        net = round(num(sheet_net_raw))
        if sheet_gross_raw != "" and num(sheet_gross_raw) > 0:
            gross_earn = round(num(sheet_gross_raw))
            # Ensure gross_earn - gross_ded = net consistency
            if abs((gross_earn - gross_ded) - net) <= 3:
                gross_earn = net + gross_ded
        else:
            gross_earn = net + gross_ded
        gross_ent = gross_earn
    else:
        # Fallback to calculated values if NetSal column not provided
        if sheet_gross_raw != "" and num(sheet_gross_raw) > 0:
            gross_earn = round(num(sheet_gross_raw))
        else:
            gross_earn = calc_gross_earn
        gross_ent = b_ent + h_ent + add_ent + conv_earn + wash_earn + leave_enc_earn + incentive_earn + extra_day_earn + ot_earn + bon_ent
        net = gross_earn - gross_ded

    return {
        "basic_ent": b_ent, "basic_earn": b_earn, "basic_arr": b_arr,
        "hra_ent": h_ent, "hra_earn": h_earn, "hra_arr": h_arr,
        "add_ent": add_ent, "add_earn": add_earn, "add_arr": add_arr,
        "conv_earn": conv_earn,
        "wash_earn": wash_earn,
        "leave_enc_earn": leave_enc_earn,
        "incentive_earn": incentive_earn,
        "extra_day_earn": extra_day_earn,
        "ot_earn": ot_earn,
        "bonus_ent": bon_ent, "bonus_earn": bon_earn, "bonus_arr": bon_arr,
        "pf_ded": pf_ded,
        "esi_ded": esi_ded,
        "pt_ded": pt_ded,
        "lwf_ded": lwf_ded,
        "gross_ent": gross_ent,
        "gross_earn": gross_earn,
        "gross_ded": gross_ded,
        "net": net
    }

def build_pdf(row, mapping, month, company=DEFAULT_COMPANY, address=DEFAULT_ADDRESS, location=""):
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
    emp_loc = (getval(row, mapping, "location") or location or "").strip()
    emp_unit = (getval(row, mapping, "unit") or emp_loc).strip()

    right_fields = [
        ("Location",         emp_loc or "Site Office"),
        ("Unit",             emp_unit or "Site Office"),
        ("Joining Date",     getval(row, mapping, "joining_date") or ""),
        ("Paid days",        f"{float(paid_d):.2f}" if paid_d and str(paid_d).replace('.', '', 1).isdigit() else paid_d),
        ("Bank Account No.", getval(row, mapping, "bank", "N/A")),
        ("PF No.",           getval(row, mapping, "pf_no")),
        ("Cost Center",      getval(row, mapping, "cost_center", company)),
    ]
    extra_day_att_val = num(getval(row, mapping, "extra_day_att", getval(row, mapping, "extra_day", 0)))
    ot_hrs_val = num(getval(row, mapping, "ot_hours", getval(row, mapping, "ot", 0)))

    attend_fields = [
        ("Present", pres_d),
        ("LWP",     f"{float(lwp_d):.2f}" if lwp_d and str(lwp_d).replace('.', '', 1).isdigit() else "0.00"),
        ("NH",      f"{float(getval(row, mapping, 'nh', '0')):.2f}" if getval(row, mapping, 'nh') and str(getval(row, mapping, 'nh')).replace('.', '', 1).isdigit() else "0.00"),
        ("WOFF",    f"{float(getval(row, mapping, 'woff', '4')):.2f}" if getval(row, mapping, 'woff') and str(getval(row, mapping, 'woff')).replace('.', '', 1).isdigit() else "4.00"),
        ("Extra Day", f"{extra_day_att_val:.2f}"),
        ("OT",      f"{ot_hrs_val:.2f}"),
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

    earnings_rows = []
    earnings_rows.append(("BASIC", fin["basic_ent"], fin["basic_earn"], fin["basic_arr"]))
    earnings_rows.append(("HRA", fin["hra_ent"], fin["hra_earn"], fin["hra_arr"]))
    if fin.get("add_earn", 0) > 0 or "add_allowance" in mapping:
        earnings_rows.append(("Additional Allowance", fin["add_ent"], fin["add_earn"], fin["add_arr"]))
    if fin.get("conv_earn", 0) > 0:
        earnings_rows.append(("Conveyance Allowance", fin["conv_earn"], fin["conv_earn"], 0))
    if fin.get("wash_earn", 0) > 0:
        earnings_rows.append(("Washing Allowance", fin["wash_earn"], fin["wash_earn"], 0))
    if fin.get("leave_enc_earn", 0) > 0:
        earnings_rows.append(("Leave Encashment", fin["leave_enc_earn"], fin["leave_enc_earn"], 0))
    if fin.get("extra_day_earn", 0) > 0:
        earnings_rows.append(("Extra Day Pay", fin["extra_day_earn"], fin["extra_day_earn"], 0))
    if fin.get("ot_earn", 0) > 0:
        earnings_rows.append(("OT", fin["ot_earn"], fin["ot_earn"], 0))
    if fin.get("incentive_earn", 0) > 0:
        earnings_rows.append(("Incentive", fin["incentive_earn"], fin["incentive_earn"], 0))
    if fin.get("bonus_earn", 0) > 0 or fin.get("bonus_ent", 0) > 0 or "bonus_entitled" in mapping:
        earnings_rows.append(("Bonus", fin["bonus_ent"], fin["bonus_earn"], fin["bonus_arr"]))

    deductions_rows = []
    if fin.get("pf_ded", 0) > 0 or "pf" in mapping:
        deductions_rows.append(("PF", fin["pf_ded"]))
    if fin.get("esi_ded", 0) > 0 or "esi" in mapping:
        deductions_rows.append(("ESI", fin["esi_ded"]))
    if fin.get("pt_ded", 0) > 0 or "prof_tax" in mapping:
        deductions_rows.append(("Prof. Tax", fin["pt_ded"]))
    if fin.get("lwf_ded", 0) > 0 or "lwf" in mapping:
        deductions_rows.append(("LWF", fin["lwf_ded"]))

    earnings_table_data = [
        [
            Paragraph("<b>Earnings</b>", table_header_style),
            Paragraph("<b>Entitled Amt.</b>", table_right_header),
            Paragraph("<b>Earned Amt.</b>", table_right_header),
            Paragraph("<b>Arrears</b>", table_right_header),
            Paragraph("<b>Deductions</b>", table_header_style),
            Paragraph("<b>Amount</b>", table_right_header),
        ]
    ]

    num_rows = max(len(earnings_rows), len(deductions_rows))
    for i in range(num_rows):
        if i < len(earnings_rows):
            e_name, e_ent, e_earn, e_arr = earnings_rows[i]
            c1 = Paragraph(e_name, normal_text)
            c2 = Paragraph(money(e_ent), cell_right)
            c3 = Paragraph(money(e_earn), cell_right)
            c4 = Paragraph(money(e_arr) if e_arr else "", cell_right)
        else:
            c1 = Paragraph("", normal_text)
            c2 = Paragraph("", cell_right)
            c3 = Paragraph("", cell_right)
            c4 = Paragraph("", cell_right)

        if i < len(deductions_rows):
            d_name, d_amt = deductions_rows[i]
            c5 = Paragraph(d_name, normal_text)
            c6 = Paragraph(money(d_amt), cell_right)
        else:
            c5 = Paragraph("", normal_text)
            c6 = Paragraph("", cell_right)

        earnings_table_data.append([c1, c2, c3, c4, c5, c6])

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
            # Smart sheet reader: scans up to 15 rows for headers and detects site banners (e.g. "Adidev - Goa")
            if ext == ".csv":
                raw_df = pd.read_csv(f, dtype=str, header=None).fillna("")
            else:
                raw_df = pd.read_excel(f, dtype=str, header=None).fillna("")

            header_row_idx = 0
            max_score = -1
            detected_site_name = ""

            for r_idx in range(min(15, len(raw_df))):
                row_vals = [norm(str(v)) for v in raw_df.iloc[r_idx] if str(v).strip()]
                if not row_vals:
                    continue

                # Check if a non-empty row before headers looks like a site / company banner
                if max_score < 2:
                    banner_cand = " ".join(str(v).strip() for v in raw_df.iloc[r_idx] if str(v).strip())
                    if len(banner_cand) >= 3 and not any(k in norm(banner_cand) for k in ["name", "code", "emp", "basic", "salary", "designation"]):
                        detected_site_name = banner_cand

                score = sum(1 for v in row_vals if any(k in v for k in [
                    "name", "code", "emp", "basic", "salary", "gross", "net",
                    "ded", "pf", "esi", "days", "pan", "dept", "desig", "leave", "hra", "bonus"
                ]))
                if score > max_score:
                    max_score = score
                    header_row_idx = r_idx

            if max_score >= 2:
                # Rename duplicate columns (e.g. Exra Day -> Exra Day and Exra Day.1)
                cols = []
                seen = {}
                for c in raw_df.iloc[header_row_idx]:
                    clean_c = str(c).strip()
                    if not clean_c or clean_c.lower().startswith("unnamed"):
                        cols.append("")
                        continue
                    if clean_c in seen:
                        seen[clean_c] += 1
                        cols.append(f"{clean_c}.{seen[clean_c]}")
                    else:
                        seen[clean_c] = 0
                        cols.append(clean_c)

                df = raw_df.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
                df.columns = cols
                df = df.loc[:, [bool(c) for c in df.columns]]
            else:
                f.seek(0)
                if ext == ".csv":
                    df = pd.read_csv(f, dtype=str).fillna("")
                else:
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
        form_location = request.form.get("location", "").strip()
        detected_loc = form_location or detect_location(detected_site_name, f.filename)
        location = detected_loc.upper() if detected_loc else ""

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
            "total_fields": len(FIELDS),
            "location": location
        }

        # Build a single JSON blob for safe embedding in a <script> tag
        # (avoids Jinja2 HTML auto-escaping of quotes in hidden input values)
        payroll_data_json = json.dumps({
            "records": records,
            "mapping": mapping,
            "month": month,
            "company": company,
            "address": address,
            "location": location
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
            location=location,
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
        location = data.get("location", "").strip()

        row = pd.Series(row_dict)
        pdf_bytes = build_pdf(row, mapping, month, company, address, location=location)

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
    location = request.form.get("location", "").strip()
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
        pdf = build_pdf(row, mapping, month, company, address, location=location)
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
            "OT": "0",
            "Extra Day": "0",
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
            "OT": "0",
            "Extra Day": "0",
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
            "OT": "0",
            "Extra Day": "0",
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
