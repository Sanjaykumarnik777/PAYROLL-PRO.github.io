import os
import sqlite3
import datetime
import zipfile
import calendar
from io import BytesIO
from functools import wraps
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

import pandas as pd
import razorpay
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, jsonify
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

app = Flask(__name__)
app.secret_key = "change-this-secret-key"

# =========================
# LOCAL PAYMENT TEST CONFIG
# =========================
RAZORPAY_KEY_ID = "rzp_test_SfKO3IFwsgnWhC"
RAZORPAY_KEY_SECRET = "F7CyUtorZhRe4q8YXtIefnv9"

razorpay_client = razorpay.Client(
    auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET)
)

# Local testing ke liye True rakho
PAYMENTS_ENABLED = False

# Local payment test ke liye False rakho
DEMO_MODE = False


FREE_EMPLOYEE_LIMIT = 10

ADMIN_USERNAMES = [
    "smarthireai5"
]

DB_NAME = "payroll_pro.db"
UPLOAD_FOLDER = "uploads"
PAYSLIP_FOLDER = "payslips"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PAYSLIP_FOLDER, exist_ok=True)


# ---------------------------
# DATABASE
# ---------------------------
def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def safe_add_column(cur, table_name, column_name, column_type):
    cur.execute(f"PRAGMA table_info({table_name})")
    existing_columns = [col[1] for col in cur.fetchall()]

    if column_name not in existing_columns:
        cur.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")

def clean_text(value, default=""):
    if pd.isna(value):
        return default
    value = str(value).strip()
    if value.lower() in ["nan", "none", "null"]:
        return default
    return value


def clean_float(value, default=0):
    try:
        if pd.isna(value) or value == "":
            return default
        return float(value)
    except Exception:
        return default

def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT NOT NULL
    )
    """)

    safe_add_column(cur, "companies", "address", "TEXT DEFAULT ''")
    safe_add_column(cur, "companies", "email", "TEXT DEFAULT ''")
    safe_add_column(cur, "companies", "phone", "TEXT DEFAULT ''")
    safe_add_column(cur, "companies", "overtime_multiplier", "REAL DEFAULT 1")
    safe_add_column(cur, "companies", "working_days_policy", "TEXT DEFAULT 'attendance'")


    cur.execute("""
    CREATE TABLE IF NOT EXISTS compliance_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL UNIQUE,

        pf_employee_rate REAL DEFAULT 12,
        pf_employer_rate REAL DEFAULT 12,
        pf_wage_ceiling REAL DEFAULT 15000,
        pf_max_deduction REAL DEFAULT 1800,

        esic_employee_rate REAL DEFAULT 0.75,
        esic_employer_rate REAL DEFAULT 3.25,
        esic_wage_limit REAL DEFAULT 21000,

        gratuity_rate REAL DEFAULT 4.81,
        bonus_rate REAL DEFAULT 8.33,

        tds_enabled INTEGER DEFAULT 0,

        salary_days_policy TEXT DEFAULT 'attendance',
        custom_salary_days REAL DEFAULT 30,
        count_weekly_off_paid INTEGER DEFAULT 1,
        count_paid_leave_paid INTEGER DEFAULT 1,
        count_holiday_paid INTEGER DEFAULT 1,
        deduct_lop INTEGER DEFAULT 1,

        festival_bonus_enabled INTEGER DEFAULT 0,
        festival_bonus_month INTEGER DEFAULT 10,

        updated_at TEXT,

        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
""")

    compliance_columns = {
    "salary_days_policy": "TEXT DEFAULT 'attendance'",
    "custom_salary_days": "REAL DEFAULT 30",
    "count_weekly_off_paid": "INTEGER DEFAULT 1",
    "count_paid_leave_paid": "INTEGER DEFAULT 1",
    "count_holiday_paid": "INTEGER DEFAULT 1",
    "deduct_lop": "INTEGER DEFAULT 1",

    "festival_bonus_enabled": "INTEGER DEFAULT 0",
    "festival_bonus_month": "INTEGER DEFAULT 10",
    "bonus_min_service_days": "INTEGER DEFAULT 30",
    "bonus_prorata_enabled": "INTEGER DEFAULT 1"
}

    cur.execute("PRAGMA table_info(compliance_settings)")
    existing_columns = [col[1] for col in cur.fetchall()]

    for column_name, column_type in compliance_columns.items():
        if column_name not in existing_columns:
           cur.execute(f"""
            ALTER TABLE compliance_settings
            ADD COLUMN {column_name} {column_type}
        """)

    safe_add_column(cur, "compliance_settings", "salary_days_policy", "TEXT DEFAULT 'attendance'")
    safe_add_column(cur, "compliance_settings", "custom_salary_days", "REAL DEFAULT 30")
    safe_add_column(cur, "compliance_settings", "count_weekly_off_paid", "INTEGER DEFAULT 1")
    safe_add_column(cur, "compliance_settings", "count_paid_leave_paid", "INTEGER DEFAULT 1")
    safe_add_column(cur, "compliance_settings", "count_holiday_paid", "INTEGER DEFAULT 1")
    safe_add_column(cur, "compliance_settings", "deduct_lop", "INTEGER DEFAULT 1")
   
    safe_add_column(cur, "compliance_settings", "festival_bonus_enabled", "INTEGER DEFAULT 0")
    safe_add_column(cur, "compliance_settings", "festival_bonus_month", "INTEGER DEFAULT 10")

    safe_add_column(cur, "compliance_settings", "bonus_min_service_days", "INTEGER DEFAULT 30")
    safe_add_column(cur, "compliance_settings", "bonus_prorata_enabled", "INTEGER DEFAULT 1")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        full_name TEXT NOT NULL,
        username TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        emp_code TEXT NOT NULL,
        employee_name TEXT NOT NULL,
        role TEXT NOT NULL,
        department TEXT,
        gender TEXT DEFAULT 'male',
        monthly_salary REAL NOT NULL,
        tax_regime TEXT DEFAULT 'new',
        other_annual_deductions REAL DEFAULT 0,
        special_allowance REAL DEFAULT 0,
        UNIQUE(company_id, emp_code),
        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
    """)

    # Extra employee details for payslip
    safe_add_column(cur, "employees", "uan_no", "TEXT DEFAULT ''")
    safe_add_column(cur, "employees", "esic_no", "TEXT DEFAULT ''")
    safe_add_column(cur, "employees", "bank_name", "TEXT DEFAULT ''")
    safe_add_column(cur, "employees", "account_no", "TEXT DEFAULT ''")
    safe_add_column(cur, "employees", "ifsc_code", "TEXT DEFAULT ''")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        emp_code TEXT NOT NULL,
        month TEXT NOT NULL,
        working_days INTEGER NOT NULL,
        present_days INTEGER NOT NULL,
        overtime_hours REAL DEFAULT 0,
        bonus REAL DEFAULT 0,
        manual_deduction REAL DEFAULT 0,
        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
    """)

    safe_add_column(cur, "attendance", "weekly_off", "REAL DEFAULT 0")
    safe_add_column(cur, "attendance", "paid_leave", "REAL DEFAULT 0")
    safe_add_column(cur, "attendance", "holiday", "REAL DEFAULT 0")
    safe_add_column(cur, "attendance", "lop_days", "REAL DEFAULT 0")
    safe_add_column(cur, "attendance", "paid_days", "REAL DEFAULT 0")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS leave_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        emp_code TEXT NOT NULL,
        leave_type TEXT NOT NULL,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        total_days REAL DEFAULT 0,
        reason TEXT,
        status TEXT DEFAULT 'Pending',
        admin_remark TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS leave_balances (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        emp_code TEXT NOT NULL,
        casual_leave REAL DEFAULT 0,
        sick_leave REAL DEFAULT 0,
        paid_leave REAL DEFAULT 0,
        used_leave REAL DEFAULT 0,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(company_id, emp_code)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS leave_policy_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL UNIQUE,
        casual_leave_limit REAL DEFAULT 6,
        sick_leave_limit REAL DEFAULT 6,
        paid_leave_limit REAL DEFAULT 12,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS payroll_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        emp_code TEXT NOT NULL,
        employee_name TEXT NOT NULL,
        role TEXT NOT NULL,
        department TEXT,
        gender TEXT,
        month TEXT NOT NULL,
        monthly_salary REAL NOT NULL,
        basic REAL NOT NULL,
        da REAL NOT NULL,
        hra REAL NOT NULL,
        special_allowance REAL NOT NULL,
        other_allowance REAL NOT NULL,
        gross REAL NOT NULL,
        esi_employee REAL NOT NULL,
        professional_tax REAL NOT NULL,
        pf_employee REAL NOT NULL,
        lwf_employee REAL NOT NULL,
        tds REAL NOT NULL,
        manual_deduction REAL NOT NULL,
        total_deductions REAL NOT NULL,
        esi_employer REAL NOT NULL,
        pf_employer REAL NOT NULL,
        gratuity REAL NOT NULL,
        bonus_ctc REAL NOT NULL,
        festival_bonus REAL NOT NULL,
        lwf_employer REAL NOT NULL,
        total_contributions REAL NOT NULL,
        net_pay REAL NOT NULL,
        monthly_ctc REAL NOT NULL,
        annual_ctc REAL NOT NULL,
        overtime_hours REAL NOT NULL,
        overtime_amount REAL NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
    """)

    safe_add_column(cur, "payroll_history", "run_id", "TEXT")
    safe_add_column(cur, "payroll_history", "is_current", "INTEGER DEFAULT 1")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS full_final_settlements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        emp_code TEXT NOT NULL,
        employee_name TEXT,
        role TEXT,
        department TEXT,
        last_working_day TEXT NOT NULL,
        settlement_month TEXT NOT NULL,

    monthly_salary REAL DEFAULT 0,
    paid_days REAL DEFAULT 0,
    earned_salary REAL DEFAULT 0,

    leave_balance REAL DEFAULT 0,
    leave_encashment REAL DEFAULT 0,
    bonus_payable REAL DEFAULT 0,
    gratuity_payable REAL DEFAULT 0,
    other_earnings REAL DEFAULT 0,

    notice_recovery REAL DEFAULT 0,
    loan_recovery REAL DEFAULT 0,
    advance_recovery REAL DEFAULT 0,
    other_deductions REAL DEFAULT 0,

    total_earnings REAL DEFAULT 0,
    total_deductions REAL DEFAULT 0,
    final_payable REAL DEFAULT 0,

    reason TEXT,
    remarks TEXT,
    created_at TEXT NOT NULL,

    FOREIGN KEY(company_id) REFERENCES companies(id)
)
""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS subscriptions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        plan_name TEXT NOT NULL,
        status TEXT NOT NULL,
        start_date TEXT,
        end_date TEXT,
        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER NOT NULL,
        amount REAL NOT NULL,
        payment_id TEXT,
        order_id TEXT,
        status TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(company_id) REFERENCES companies(id)
    )
    """)

    safe_add_column(cur, "payments", "user_id", "INTEGER")

    conn.commit()
    conn.close()


def ensure_leave_tables():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS leave_policy (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            casual_leave_limit REAL DEFAULT 6,
            sick_leave_limit REAL DEFAULT 6,
            paid_leave_limit REAL DEFAULT 12,
            created_at TEXT,
            UNIQUE(company_id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS leave_balances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            emp_code TEXT NOT NULL,
            casual_leave REAL DEFAULT 6,
            sick_leave REAL DEFAULT 6,
            paid_leave REAL DEFAULT 12,
            used_leave REAL DEFAULT 0,
            UNIQUE(company_id, emp_code)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS leave_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            emp_code TEXT NOT NULL,
            leave_type TEXT NOT NULL,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            total_days REAL DEFAULT 0,
            reason TEXT,
            status TEXT DEFAULT 'Pending',
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()


def add_leave_payroll_columns():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("PRAGMA table_info(payroll_history)")
    existing_columns = [col["name"] for col in cur.fetchall()]

    new_columns = {
        "paid_leave_days": "REAL DEFAULT 0",
        "lwp_days": "REAL DEFAULT 0",
        "lwp_deduction": "REAL DEFAULT 0",
        "payable_days": "REAL DEFAULT 0"
    }

    for column_name, column_type in new_columns.items():
        if column_name not in existing_columns:
            cur.execute(f"""
                ALTER TABLE payroll_history
                ADD COLUMN {column_name} {column_type}
            """)

    conn.commit()
    conn.close()


def add_payment_order_id_column():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("PRAGMA table_info(payments)")
    existing_columns = [col["name"] for col in cur.fetchall()]

    if "order_id" not in existing_columns:
        cur.execute("""
            ALTER TABLE payments
            ADD COLUMN order_id TEXT
        """)

    conn.commit()
    conn.close()


# ---------------------------
# HELPERS
# ---------------------------
def rupee(value):
    return int(round(float(value or 0)))


def money_str(value):
    return str(int(round(float(value or 0))))


def month_only(payroll_month):
    if isinstance(payroll_month, str) and "-" in payroll_month:
        return payroll_month.split("-")[1]
    return str(payroll_month)


def current_company_id():
    return session.get("company_id")


def is_admin_user():
    user_id = session.get("user_id")

    if not user_id:
        return False

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT username
        FROM users
        WHERE id = ?
    """, (user_id,))

    user = cur.fetchone()
    conn.close()

    if not user:
        return False

    username = str(user["username"] or "").strip().lower()

    # ADMIN_USERNAMES list se admin check
    admin_usernames = [str(admin_username).strip().lower() for admin_username in ADMIN_USERNAMES]

    # Local / owner testing ke liye direct allowed users
    owner_usernames = [
        "sai.enterprises7310@gmail.com",
        "admin",
        "sanjay",
        "mansi international"
    ]

    return username in admin_usernames or username in owner_usernames


def create_error_report(row_errors, filename="upload_errors.xlsx"):
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    df_errors = pd.DataFrame({"Error": row_errors})
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    df_errors.to_excel(file_path, index=False)
    return file_path


def ensure_subscription_valid():
    conn = get_db()
    cur = conn.cursor()
    today = datetime.datetime.now().strftime("%Y-%m-%d")

    cur.execute("""
        UPDATE subscriptions
        SET status = 'expired'
        WHERE status = 'active'
        AND end_date IS NOT NULL
        AND end_date < ?
    """, (today,))

    conn.commit()
    conn.close()


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        ensure_subscription_valid()
        return fn(*args, **kwargs)
    return wrapper


def get_active_plan():
    company_id = current_company_id()

    default_free_plan = {
        "is_pro": False,
        "plan": "FREE",
        "status": "free",
        "start_date": "-",
        "end_date": "-"
    }

    if not company_id:
        return default_free_plan

    conn = None

    try:
        conn = get_db()
        cur = conn.cursor()

        today = datetime.datetime.now().strftime("%Y-%m-%d")

        cur.execute("""
            SELECT 
                plan_name,
                status,
                start_date,
                end_date
            FROM subscriptions
            WHERE company_id = ?
              AND LOWER(COALESCE(status, '')) = 'active'
              AND end_date IS NOT NULL
              AND TRIM(end_date) != ''
              AND date(end_date) >= date(?)
            ORDER BY date(end_date) DESC, id DESC
            LIMIT 1
        """, (company_id, today))

        sub = cur.fetchone()

        if not sub:
            return default_free_plan

        plan_name = str(sub["plan_name"] or "PRO").strip().upper()

        # Safety: only real paid/pro plans should unlock PRO features.
        # If accidentally FREE/LIFETIME FREE is saved as active, do not unlock.
        free_plan_names = ["FREE", "FREE PLAN", "LIFETIME FREE", "BASIC"]

        if plan_name in free_plan_names:
            return default_free_plan

        return {
            "is_pro": True,
            "plan": plan_name,
            "status": sub["status"] or "active",
            "start_date": sub["start_date"] or "-",
            "end_date": sub["end_date"] or "-"
        }

    except Exception:
        # Subscription check fail hua to safe side: FREE access.
        return default_free_plan

    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def is_campaign_free_mode():
    """
    Launch campaign mode.
    When ON, all PRO/download features are temporarily free for users.
    Later, turn this OFF from Render environment variable.
    """
    return os.environ.get("CAMPAIGN_FREE_MODE", "off").lower() == "on"

@app.context_processor
def inject_layout_plan():
    try:
        if session.get("user_id"):
            active_plan = get_active_plan()
            return {
                "layout_active_plan": active_plan
            }
    except Exception:
        pass

    return {
        "layout_active_plan": {
            "is_pro": False,
            "plan": "FREE",
            "status": "free",
            "start_date": "-",
            "end_date": "-"
        }
    }

def get_compliance_settings(company_id):
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM compliance_settings
        WHERE company_id = ?
    """, (company_id,))

    settings = cur.fetchone()

    if not settings:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cur.execute("""
            INSERT INTO compliance_settings (
                company_id,

                pf_employee_rate,
                pf_employer_rate,
                pf_wage_ceiling,
                pf_max_deduction,

                esic_employee_rate,
                esic_employer_rate,
                esic_wage_limit,

                gratuity_rate,
                bonus_rate,
                tds_enabled,

                salary_days_policy,
                custom_salary_days,
                count_weekly_off_paid,
                count_paid_leave_paid,
                count_holiday_paid,
                deduct_lop,

                festival_bonus_enabled,
                festival_bonus_month,
                bonus_min_service_days,
                bonus_prorata_enabled,

                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            company_id,

            12,
            12,
            15000,
            1800,

            0.75,
            3.25,
            21000,

            4.81,
            8.33,
            0,

            "attendance",
            30,
            1,
            1,
            1,
            1,

            0,
            10,
            30,
            1,

            now
        ))

        conn.commit()

        cur.execute("""
            SELECT *
            FROM compliance_settings
            WHERE company_id = ?
        """, (company_id,))

        settings = cur.fetchone()

    conn.close()
    return settings


def is_pro_user():
    if is_campaign_free_mode():
        return True

    if is_admin_user():
        return True

    active_plan = get_active_plan()

    if active_plan and active_plan.get("is_pro") is True:
        return True

    return False


def require_pro_feature(message="Upgrade to PRO to use this feature."):
    if is_campaign_free_mode():
        return True

    if is_admin_user():
        return True

    active_plan = get_active_plan()

    if active_plan and active_plan.get("is_pro") is True:
        return True

    flash(message, "warning")
    return False


def get_employee_limit():
    # Admin/demo owner unlimited employees
    if is_admin_user():
        return None

    active_plan = get_active_plan()

    if active_plan and active_plan.get("is_pro") is True:
        return None

    return FREE_EMPLOYEE_LIMIT


def can_add_employee():
    # Admin/demo owner ko limit nahi lagegi
    if is_admin_user():
        return True, ""

    active_plan = get_active_plan()

    if active_plan and active_plan.get("is_pro") is True:
        return True, ""

    company_id = current_company_id()

    if not company_id:
        return False, "Company not found. Please login again."

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT COUNT(*) AS count
        FROM employees
        WHERE company_id = ?
    """, (company_id,))

    result = cur.fetchone()
    employee_count = result["count"] if result else 0

    conn.close()

    if employee_count >= FREE_EMPLOYEE_LIMIT:
        return False, (
            f"Free plan allows up to {FREE_EMPLOYEE_LIMIT} employees only. "
            "Please upgrade to PRO to add unlimited employees."
        )

    return True, ""


# ---------------------------
# COMPLIANCE RULES
# ---------------------------
def calculate_professional_tax_maharashtra(gross_salary, gender, payroll_month):
    gross_salary = float(gross_salary or 0)
    gender = str(gender or "male").strip().lower()
    month = month_only(payroll_month)

    # Female employee Maharashtra PT rule
    if gender == "female":
        if gross_salary <= 25000:
            return 0
        return 300 if month == "02" else 200

    # Male employee Maharashtra PT rule
    if gross_salary <= 7500:
        return 0

    if gross_salary <= 10000:
        return 175

    # Male salary above 10000
    return 300 if month == "02" else 200


def calculate_lwf_maharashtra(payroll_month_mm):
    payroll_month_mm = month_only(payroll_month_mm)
    if payroll_month_mm in ["06", "12"]:
        return {"employee": 25.0, "employer": 75.0}
    return {"employee": 0.0, "employer": 0.0}


def calculate_bonus_logic(basic, payroll_month, bonus_rate=0.0833):
    basic = float(basic or 0)

    # Bonus is part of CTC as per compliance setting
    bonus_ctc = rupee(basic * bonus_rate)

    # Regular monthly salary me bonus payout nahi.
    # Festival/Diwali payout separately control kar sakte ho.
    festival_bonus = 0

    return bonus_ctc, festival_bonus


@app.route("/compliance-settings", methods=["GET", "POST"])
@login_required
def compliance_settings():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    # Ensure default settings exist for this company
    get_compliance_settings(company_id)

    def to_float(name, default=0):
        try:
            value = request.form.get(name, default)
            if value is None or str(value).strip() == "":
                return float(default)
            return float(value)
        except Exception:
            return float(default)

    def to_int(name, default=0):
        try:
            value = request.form.get(name, default)
            if value is None or str(value).strip() == "":
                return int(default)
            return int(float(value))
        except Exception:
            return int(default)

    def checkbox_value(name):
        return 1 if request.form.get(name) in ["1", "on", "true", "True", "yes", "Yes"] else 0

    def clamp(value, minimum, maximum, default):
        try:
            value = float(value)
            if value < minimum or value > maximum:
                return default
            return value
        except Exception:
            return default

    if request.method == "POST":
        conn = get_db()
        cur = conn.cursor()

        try:
            # PF settings
            pf_employee_rate = to_float("pf_employee_rate", 12)
            pf_employer_rate = to_float("pf_employer_rate", 12)
            pf_wage_ceiling = to_float("pf_wage_ceiling", 15000)
            pf_max_deduction = to_float("pf_max_deduction", 1800)

            # ESIC settings
            esic_employee_rate = to_float("esic_employee_rate", 0.75)
            esic_employer_rate = to_float("esic_employer_rate", 3.25)
            esic_wage_limit = to_float("esic_wage_limit", 21000)

            # Other statutory settings
            gratuity_rate = to_float("gratuity_rate", 4.81)
            bonus_rate = to_float("bonus_rate", 8.33)
            tds_enabled = checkbox_value("tds_enabled")

            # Payroll days policy
            salary_days_policy = request.form.get("salary_days_policy", "attendance").strip()
            allowed_policies = ["attendance", "fixed_26", "fixed_30", "calendar", "custom"]

            if salary_days_policy not in allowed_policies:
                salary_days_policy = "attendance"

            custom_salary_days = to_float("custom_salary_days", 30)

            count_weekly_off_paid = checkbox_value("count_weekly_off_paid")
            count_paid_leave_paid = checkbox_value("count_paid_leave_paid")
            count_holiday_paid = checkbox_value("count_holiday_paid")
            deduct_lop = checkbox_value("deduct_lop")

            # Festival bonus settings
            festival_bonus_enabled = checkbox_value("festival_bonus_enabled")
            festival_bonus_month = to_int("festival_bonus_month", 10)
            bonus_min_service_days = to_int("bonus_min_service_days", 30)
            bonus_prorata_enabled = checkbox_value("bonus_prorata_enabled")

            errors = []

            # Validation: percentage fields
            rate_fields = {
                "PF Employee Rate": pf_employee_rate,
                "PF Employer Rate": pf_employer_rate,
                "ESIC Employee Rate": esic_employee_rate,
                "ESIC Employer Rate": esic_employer_rate,
                "Gratuity Rate": gratuity_rate,
                "Bonus Rate": bonus_rate
            }

            for label, value in rate_fields.items():
                if value < 0:
                    errors.append(f"{label} cannot be negative.")

                if value > 100:
                    errors.append(f"{label} cannot be more than 100%.")

            # Validation: amount / limit fields
            if pf_wage_ceiling < 0:
                errors.append("PF Wage Ceiling cannot be negative.")

            if pf_wage_ceiling > 10000000:
                errors.append("PF Wage Ceiling amount is too high. Please check.")

            if pf_max_deduction < 0:
                errors.append("PF Max Deduction cannot be negative.")

            if pf_max_deduction > 10000000:
                errors.append("PF Max Deduction amount is too high. Please check.")

            if esic_wage_limit < 0:
                errors.append("ESIC Wage Limit cannot be negative.")

            if esic_wage_limit > 10000000:
                errors.append("ESIC Wage Limit amount is too high. Please check.")

            # Validation: salary days
            if salary_days_policy == "custom":
                if custom_salary_days <= 0:
                    errors.append("Custom Salary Days must be greater than 0.")

                if custom_salary_days > 31:
                    errors.append("Custom Salary Days cannot be more than 31.")
            else:
                if custom_salary_days <= 0 or custom_salary_days > 31:
                    custom_salary_days = 30

            # Validation: bonus month and service days
            if festival_bonus_month < 1 or festival_bonus_month > 12:
                festival_bonus_month = 10

            if bonus_min_service_days < 0:
                bonus_min_service_days = 30

            if bonus_min_service_days > 3650:
                errors.append("Bonus Minimum Service Days is too high. Please check.")

            if errors:
                flash(" ".join(errors), "danger")
                return redirect(url_for("compliance_settings"))

            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cur.execute("""
                INSERT INTO compliance_settings (
                    company_id,

                    pf_employee_rate,
                    pf_employer_rate,
                    pf_wage_ceiling,
                    pf_max_deduction,

                    esic_employee_rate,
                    esic_employer_rate,
                    esic_wage_limit,

                    gratuity_rate,
                    bonus_rate,
                    tds_enabled,

                    salary_days_policy,
                    custom_salary_days,
                    count_weekly_off_paid,
                    count_paid_leave_paid,
                    count_holiday_paid,
                    deduct_lop,

                    festival_bonus_enabled,
                    festival_bonus_month,
                    bonus_min_service_days,
                    bonus_prorata_enabled,

                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(company_id) DO UPDATE SET
                    pf_employee_rate = excluded.pf_employee_rate,
                    pf_employer_rate = excluded.pf_employer_rate,
                    pf_wage_ceiling = excluded.pf_wage_ceiling,
                    pf_max_deduction = excluded.pf_max_deduction,

                    esic_employee_rate = excluded.esic_employee_rate,
                    esic_employer_rate = excluded.esic_employer_rate,
                    esic_wage_limit = excluded.esic_wage_limit,

                    gratuity_rate = excluded.gratuity_rate,
                    bonus_rate = excluded.bonus_rate,
                    tds_enabled = excluded.tds_enabled,

                    salary_days_policy = excluded.salary_days_policy,
                    custom_salary_days = excluded.custom_salary_days,
                    count_weekly_off_paid = excluded.count_weekly_off_paid,
                    count_paid_leave_paid = excluded.count_paid_leave_paid,
                    count_holiday_paid = excluded.count_holiday_paid,
                    deduct_lop = excluded.deduct_lop,

                    festival_bonus_enabled = excluded.festival_bonus_enabled,
                    festival_bonus_month = excluded.festival_bonus_month,
                    bonus_min_service_days = excluded.bonus_min_service_days,
                    bonus_prorata_enabled = excluded.bonus_prorata_enabled,

                    updated_at = excluded.updated_at
            """, (
                company_id,

                pf_employee_rate,
                pf_employer_rate,
                pf_wage_ceiling,
                pf_max_deduction,

                esic_employee_rate,
                esic_employer_rate,
                esic_wage_limit,

                gratuity_rate,
                bonus_rate,
                tds_enabled,

                salary_days_policy,
                custom_salary_days,
                count_weekly_off_paid,
                count_paid_leave_paid,
                count_holiday_paid,
                deduct_lop,

                festival_bonus_enabled,
                festival_bonus_month,
                bonus_min_service_days,
                bonus_prorata_enabled,

                now
            ))

            conn.commit()
            flash("Compliance, payroll policy and bonus settings updated successfully.", "success")
            return redirect(url_for("compliance_settings"))

        except Exception as e:
            conn.rollback()
            flash(f"Error saving compliance settings: {str(e)}", "danger")
            return redirect(url_for("compliance_settings"))

        finally:
            conn.close()

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT *
            FROM compliance_settings
            WHERE company_id = ?
            LIMIT 1
        """, (company_id,))

        settings = cur.fetchone()

    except Exception as e:
        flash(f"Error loading compliance settings: {str(e)}", "danger")
        settings = None

    finally:
        conn.close()

    return render_template(
        "compliance_settings.html",
        settings=settings
    )


# ---------------------------
# AUTH
# ---------------------------
@app.route("/home")
def home():
    if session.get("user_id"):
        return redirect(url_for("dashboard"))

    return render_template("landing.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        company_name = request.form["company_name"].strip()
        company_address = request.form.get("company_address", "").strip()
        company_email = request.form.get("company_email", "").strip()
        company_phone = request.form.get("company_phone", "").strip()

        full_name = request.form["full_name"].strip()
        username = request.form["username"].strip()
        password = request.form["password"].strip()

        conn = get_db()
        cur = conn.cursor()

        try:
            cur.execute("""
                INSERT INTO companies (company_name, address, email, phone)
                VALUES (?, ?, ?, ?)
            """, (company_name, company_address, company_email, company_phone))

            company_id = cur.lastrowid

            cur.execute("""
                INSERT INTO users (company_id, full_name, username, password_hash)
                VALUES (?, ?, ?, ?)
            """, (company_id, full_name, username, generate_password_hash(password)))

            conn.commit()
            flash("Registration successful. Please login.")
            return redirect(url_for("login"))

        except Exception as e:
            flash(f"Registration failed: {e}")

        finally:
            conn.close()

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()

        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username = ?", (username,))
        user = cur.fetchone()
        conn.close()

        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["company_id"] = user["company_id"]
            session["username"] = user["username"]
            return redirect(url_for("dashboard"))

        flash("Invalid username or password.")

    return render_template("login.html")


@app.route("/company-profile", methods=["GET", "POST"])
@login_required
def company_profile():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    def clean_text(value, default=""):
        value = str(value or "").strip()
        return value if value else default

    def valid_email(value):
        value = clean_text(value)
        if not value:
            return True

        if "@" not in value or "." not in value:
            return False

        return True

    def clean_phone(value):
        # Keep phone as text so leading zero is not lost.
        value = str(value or "").strip()
        return value

    conn = get_db()
    cur = conn.cursor()

    if request.method == "POST":
        try:
            company_name = clean_text(request.form.get("company_name"))
            address = clean_text(request.form.get("address"))
            email = clean_text(request.form.get("email"))
            phone = clean_phone(request.form.get("phone"))

            try:
                overtime_multiplier = float(request.form.get("overtime_multiplier", 1) or 1)
            except Exception:
                overtime_multiplier = 1

            if overtime_multiplier not in [1, 2]:
                overtime_multiplier = 1

            working_days_policy = clean_text(
                request.form.get("working_days_policy", "attendance"),
                "attendance"
            )

            allowed_working_days_policies = [
                "attendance",
                "fixed_26",
                "fixed_30"
            ]

            if working_days_policy not in allowed_working_days_policies:
                working_days_policy = "attendance"

            errors = []

            if not company_name:
                errors.append("Company name is required.")

            if len(company_name) > 150:
                errors.append("Company name is too long.")

            if len(address) > 500:
                errors.append("Company address is too long.")

            if not valid_email(email):
                errors.append("Please enter a valid company email.")

            if len(email) > 150:
                errors.append("Company email is too long.")

            if len(phone) > 30:
                errors.append("Company phone number is too long.")

            if errors:
                flash(" ".join(errors), "danger")
                return redirect(url_for("company_profile"))

            cur.execute("""
                UPDATE companies
                SET company_name = ?,
                    address = ?,
                    email = ?,
                    phone = ?,
                    overtime_multiplier = ?,
                    working_days_policy = ?
                WHERE id = ?
            """, (
                company_name,
                address,
                email,
                phone,
                overtime_multiplier,
                working_days_policy,
                company_id
            ))

            conn.commit()

            flash("Company profile updated successfully.", "success")
            return redirect(url_for("company_profile"))

        except Exception as e:
            conn.rollback()
            flash(f"Error updating company profile: {str(e)}", "danger")
            return redirect(url_for("company_profile"))

        finally:
            conn.close()

    try:
        cur.execute("""
            SELECT 
                company_name,
                address,
                email,
                phone,
                COALESCE(overtime_multiplier, 1) AS overtime_multiplier,
                COALESCE(working_days_policy, 'attendance') AS working_days_policy
            FROM companies
            WHERE id = ?
            LIMIT 1
        """, (company_id,))

        company = cur.fetchone()

        if not company:
            flash("Company profile not found. Please login again.", "danger")
            return redirect(url_for("login"))

    except Exception as e:
        flash(f"Error loading company profile: {str(e)}", "danger")
        return redirect(url_for("dashboard"))

    finally:
        conn.close()

    return render_template("company_profile.html", company=company)


# ---------------------------
# DASHBOARD / PAYMENT
# ---------------------------
@app.route("/")
@login_required
def dashboard():
    company_id = current_company_id()
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT company_name
        FROM companies
        WHERE id = ?
    """, (company_id,))
    company = cur.fetchone()

    if not company:
        conn.close()
        session.clear()
        flash("Company record not found. Please login again.")
        return redirect(url_for("login"))

    cur.execute("""
        SELECT COUNT(*) AS count
        FROM employees
        WHERE company_id = ?
    """, (company_id,))
    employee_count = cur.fetchone()["count"] or 0

    cur.execute("""
        SELECT COUNT(*) AS count
        FROM attendance
        WHERE company_id = ?
    """, (company_id,))
    attendance_count = cur.fetchone()["count"] or 0

    cur.execute("""
        SELECT COUNT(*) AS pending_leaves
        FROM leave_requests
        WHERE company_id = ?
          AND status = 'Pending'
    """, (company_id,))
    pending_leaves = cur.fetchone()["pending_leaves"] or 0

    # Latest payroll month
    cur.execute("""
        SELECT month
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
        GROUP BY month
        ORDER BY month DESC
        LIMIT 1
    """, (company_id,))
    latest_month_row = cur.fetchone()
    latest_payroll_month = latest_month_row["month"] if latest_month_row else "-"

    # Payroll count - latest/current records only
    cur.execute("""
        SELECT COUNT(*) AS count
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
    """, (company_id,))
    payroll_count = cur.fetchone()["count"] or 0

    # Overall current payroll totals
    cur.execute("""
        SELECT 
            COALESCE(SUM(gross), 0) AS total_gross,
            COALESCE(SUM(overtime_amount), 0) AS total_overtime_amount,
            COALESCE(SUM(festival_bonus), 0) AS total_festival_bonus,
            COALESCE(SUM(bonus_ctc), 0) AS total_bonus_ctc,

            COALESCE(SUM(total_deductions), 0) AS total_deductions,
            COALESCE(SUM(net_pay), 0) AS total_net_pay,

            COALESCE(SUM(pf_employer), 0) AS total_pf_employer,
            COALESCE(SUM(esi_employer), 0) AS total_esi_employer,
            COALESCE(SUM(gratuity), 0) AS total_gratuity,
            COALESCE(SUM(lwf_employer), 0) AS total_lwf_employer,

            COALESCE(SUM(monthly_ctc), 0) AS total_monthly_ctc,
            COALESCE(SUM(annual_ctc), 0) AS total_annual_ctc
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
    """, (company_id,))
    payroll_totals = cur.fetchone()

    total_gross = round(float(payroll_totals["total_gross"] or 0))
    total_overtime_amount = round(float(payroll_totals["total_overtime_amount"] or 0))
    total_festival_bonus = round(float(payroll_totals["total_festival_bonus"] or 0))
    total_bonus_ctc = round(float(payroll_totals["total_bonus_ctc"] or 0))

    total_deductions = round(float(payroll_totals["total_deductions"] or 0))
    total_net_pay = round(float(payroll_totals["total_net_pay"] or 0))

    total_pf_employer = round(float(payroll_totals["total_pf_employer"] or 0))
    total_esi_employer = round(float(payroll_totals["total_esi_employer"] or 0))
    total_gratuity = round(float(payroll_totals["total_gratuity"] or 0))
    total_lwf_employer = round(float(payroll_totals["total_lwf_employer"] or 0))

    total_employer_cost = round(
        total_pf_employer
        + total_esi_employer
        + total_gratuity
        + total_lwf_employer
    )

    total_monthly_ctc = round(float(payroll_totals["total_monthly_ctc"] or 0))
    total_annual_ctc = round(float(payroll_totals["total_annual_ctc"] or 0))

    # Latest month payroll totals
    latest_month_gross = 0
    latest_month_net_pay = 0
    latest_month_ctc = 0
    latest_month_employees = 0

    if latest_payroll_month != "-":
        cur.execute("""
            SELECT
                COUNT(*) AS employees,
                COALESCE(SUM(gross), 0) AS gross,
                COALESCE(SUM(net_pay), 0) AS net_pay,
                COALESCE(SUM(monthly_ctc), 0) AS monthly_ctc
            FROM payroll_history
            WHERE company_id = ?
              AND month = ?
              AND is_current = 1
        """, (company_id, latest_payroll_month))

        latest_totals = cur.fetchone()

        latest_month_employees = latest_totals["employees"] or 0
        latest_month_gross = round(float(latest_totals["gross"] or 0))
        latest_month_net_pay = round(float(latest_totals["net_pay"] or 0))
        latest_month_ctc = round(float(latest_totals["monthly_ctc"] or 0))

    # Subscription / active plan
    active_plan = get_active_plan()

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
        ORDER BY id DESC
        LIMIT 1
    """, (company_id,))
    subscription = cur.fetchone()

    if subscription:
        plan_name = subscription["plan_name"] or active_plan["plan"]
        subscription_status = subscription["status"] or "active"
        subscription_end_date = subscription["end_date"] or active_plan["end_date"] or "-"
    else:
        plan_name = active_plan["plan"]
        subscription_status = active_plan.get("status", "free")
        subscription_end_date = active_plan["end_date"] or "-"

    # Chart data - last 12 payroll months
    cur.execute("""
        SELECT 
            month,
            COALESCE(SUM(gross), 0) AS gross,
            COALESCE(SUM(net_pay), 0) AS net_pay,
            COALESCE(SUM(total_deductions), 0) AS deductions,
            COALESCE(SUM(monthly_ctc), 0) AS monthly_ctc
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
        GROUP BY month
        ORDER BY month ASC
        LIMIT 12
    """, (company_id,))
    chart_rows = cur.fetchall()

    chart_labels = []
    chart_gross = []
    chart_net_pay = []
    chart_deductions = []
    chart_monthly_ctc = []

    for row in chart_rows:
        chart_labels.append(row["month"])
        chart_gross.append(round(float(row["gross"] or 0)))
        chart_net_pay.append(round(float(row["net_pay"] or 0)))
        chart_deductions.append(round(float(row["deductions"] or 0)))
        chart_monthly_ctc.append(round(float(row["monthly_ctc"] or 0)))

    conn.close()

    return render_template(
        "dashboard.html",
        company_name=company["company_name"],

        employee_count=employee_count,
        attendance_count=attendance_count,
        payroll_count=payroll_count,
        pending_leaves=pending_leaves,

        latest_payroll_month=latest_payroll_month,
        latest_month_employees=latest_month_employees,
        latest_month_gross=latest_month_gross,
        latest_month_net_pay=latest_month_net_pay,
        latest_month_ctc=latest_month_ctc,

        total_gross=total_gross,
        total_overtime_amount=total_overtime_amount,
        total_festival_bonus=total_festival_bonus,
        total_bonus_ctc=total_bonus_ctc,

        total_deductions=total_deductions,
        total_net_pay=total_net_pay,

        total_pf_employer=total_pf_employer,
        total_esi_employer=total_esi_employer,
        total_gratuity=total_gratuity,
        total_lwf_employer=total_lwf_employer,
        total_employer_cost=total_employer_cost,

        total_monthly_ctc=total_monthly_ctc,
        total_annual_ctc=total_annual_ctc,

        chart_labels=chart_labels,
        chart_gross=chart_gross,
        chart_net_pay=chart_net_pay,
        chart_deductions=chart_deductions,
        chart_monthly_ctc=chart_monthly_ctc,

        active_plan=active_plan,
        plan_name=plan_name,
        subscription_status=subscription_status,
        subscription_end_date=subscription_end_date,

        now=datetime.datetime.now()
    )


@app.route("/pricing")
@login_required
def pricing():
    company_id = current_company_id()

    plans = {
        "monthly": {
            "plan_id": "monthly",
            "name": "PRO Monthly",
            "price": 999,
            "duration_days": 30,
            "label": "Valid for 30 days",
            "button_text": "Start Monthly PRO - ₹999"
        },
        "yearly": {
            "plan_id": "yearly",
            "name": "PRO Yearly",
            "price": 9999,
            "duration_days": 365,
            "label": "Valid for 365 days",
            "button_text": "Start Yearly PRO - ₹9999",
            "saving_text": "Save ₹1,989"
        }
    }

    active_plan = get_active_plan()

    conn = get_db()
    cur = conn.cursor()

    today = datetime.datetime.now().strftime("%Y-%m-%d")

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
          AND date(end_date) >= date(?)
        ORDER BY date(end_date) DESC, id DESC
        LIMIT 1
    """, (company_id, today))

    active_subscription = cur.fetchone()
    conn.close()

    return render_template(
        "pricing.html",
        razorpay_key_id=RAZORPAY_KEY_ID,
        plans=plans,
        active_subscription=active_subscription,
        active_plan=active_plan,
        payments_enabled=PAYMENTS_ENABLED
    )


@app.route("/create-order", methods=["POST"])
@login_required
def create_order():
    data = request.get_json()
    amount = int(data.get("amount", 0)) * 100

    if amount <= 0:
        return jsonify({"status": "failed", "message": "Invalid amount"}), 400

    order = razorpay_client.order.create({
        "amount": amount,
        "currency": "INR",
        "payment_capture": 1
    })
    return jsonify(order)


@app.route("/payment-success", methods=["POST"])
@login_required
def payment_success():
    if not PAYMENTS_ENABLED:
        return "Payments are currently disabled. Free demo plan is active.", 403

    user_id = session.get("user_id")
    company_id = current_company_id()

    if not company_id:
        return "Company not found. Please login again.", 400

    payment_id = request.form.get("razorpay_payment_id", "").strip()
    order_id = request.form.get("razorpay_order_id", "").strip()
    signature = request.form.get("razorpay_signature", "").strip()

    plan_id = request.form.get("plan_id", "monthly").strip().lower()

    plans = {
        "monthly": {
            "plan_name": "PRO Monthly",
            "amount": 999,
            "duration_days": 30
        },
        "yearly": {
            "plan_name": "PRO Yearly",
            "amount": 9999,
            "duration_days": 365
        }
    }

    if plan_id not in plans:
        return "Invalid plan selected", 400

    if not payment_id:
        return "Payment ID missing", 400

    selected_plan = plans[plan_id]

    amount = selected_plan["amount"]
    plan_name = selected_plan["plan_name"]
    duration_days = selected_plan["duration_days"]

    # Razorpay payment verification
    try:
        payment_details = razorpay_client.payment.fetch(payment_id)

        if not payment_details:
            return "Unable to verify payment.", 400

        razorpay_status = payment_details.get("status", "")
        razorpay_amount = int(payment_details.get("amount", 0))
        expected_amount = int(amount * 100)

        if razorpay_status not in ["authorized", "captured"]:
            return f"Payment not successful. Current status: {razorpay_status}", 400

        if razorpay_amount != expected_amount:
            return "Payment amount mismatch.", 400

        # If payment is only authorized, capture it
        if razorpay_status == "authorized":
            razorpay_client.payment.capture(payment_id, expected_amount)

    except Exception as verify_error:
        return f"Payment verification failed: {str(verify_error)}", 400

    start_date = datetime.datetime.now()
    end_date = start_date + datetime.timedelta(days=duration_days)

    conn = get_db()
    cur = conn.cursor()

    try:
        # Duplicate payment protection
        cur.execute("""
            SELECT id
            FROM payments
            WHERE company_id = ?
              AND payment_id = ?
        """, (company_id, payment_id))

        existing_payment = cur.fetchone()

        if existing_payment:
            conn.close()
            return "Payment already recorded", 200

        # Deactivate old subscriptions
        cur.execute("""
            UPDATE subscriptions
            SET status = 'inactive'
            WHERE company_id = ?
              AND status = 'active'
        """, (company_id,))

        # Activate new subscription
        cur.execute("""
            INSERT INTO subscriptions (
                company_id,
                plan_name,
                status,
                start_date,
                end_date
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            company_id,
            plan_name,
            "active",
            start_date.strftime("%Y-%m-%d"),
            end_date.strftime("%Y-%m-%d")
        ))

        # Save payment history
        cur.execute("""
            INSERT INTO payments (
                company_id,
                user_id,
                amount,
                payment_id,
                order_id,
                status,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            company_id,
            user_id,
            amount,
            payment_id,
            order_id,
            "success",
            start_date.strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()

        return "success", 200

    except Exception as e:
        conn.rollback()
        return f"Payment activation failed: {str(e)}", 500

    finally:
        conn.close()


@app.route("/payments")
@login_required
def payments():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            id,
            amount,
            payment_id,
            order_id,
            status,
            created_at
        FROM payments
        WHERE company_id = ?
        ORDER BY id DESC
    """, (company_id,))

    data = cur.fetchall()

    total_payments = len(data)

    successful_payments = 0
    pending_payments = 0
    failed_payments = 0
    total_paid_amount = 0

    for row in data:
        status = str(row["status"] or "").lower()
        amount = float(row["amount"] or 0)

        if status in ["paid", "success", "successful", "captured"]:
            successful_payments += 1
            total_paid_amount += amount
        elif status in ["pending", "created", "authorized"]:
            pending_payments += 1
        else:
            failed_payments += 1

    total_paid_amount = round(total_paid_amount)

    active_plan = get_active_plan()

    today = datetime.datetime.now().strftime("%Y-%m-%d")

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
          AND date(end_date) >= date(?)
        ORDER BY date(end_date) DESC, id DESC
        LIMIT 1
    """, (company_id, today))

    active_subscription = cur.fetchone()

    conn.close()

    return render_template(
    "payments.html",
    data=data,
    total_paid_amount=total_paid_amount,
    total_payments=total_payments,
    successful_payments=successful_payments,
    pending_payments=pending_payments,
    failed_payments=failed_payments,
    active_plan=active_plan,
    campaign_free_mode=is_campaign_free_mode()
)


# ---------------------------
# EMPLOYEES
# ---------------------------
@app.route("/upload-employees", methods=["GET", "POST"])
@login_required
def upload_employees():
    if request.method == "POST":
        company_id = current_company_id()

        if not company_id:
            flash("Company not found. Please login again.", "danger")
            return redirect(url_for("login"))

        if "file" not in request.files:
            flash("Please select a file.", "warning")
            return redirect(url_for("upload_employees"))

        file = request.files["file"]

        if not file or file.filename == "":
            flash("Please select a file.", "warning")
            return redirect(url_for("upload_employees"))

        filename = file.filename.lower()

        if not (filename.endswith(".csv") or filename.endswith(".xlsx")):
            flash("Only CSV or Excel (.xlsx) file allowed.", "danger")
            return redirect(url_for("upload_employees"))

        path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
        file.save(path)

        conn = None

        try:
            if filename.endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path, engine="openpyxl")

            # Clean column names
            df.columns = [str(col).strip().lower() for col in df.columns]

            required_columns = [
                "emp_code",
                "employee_name",
                "role",
                "monthly_salary"
            ]

            # Required column validation - helper function ki zaroorat nahi
            missing_columns = [
                col for col in required_columns
                if col not in df.columns
            ]

            if missing_columns:
                session["error_report"] = create_error_report(
                    [f"Missing required column: {col}" for col in missing_columns],
                    "employee_upload_errors.xlsx"
                )
                flash("Upload failed. Required columns are missing. Please download the error report.", "danger")
                return redirect(url_for("upload_employees"))

            row_errors = []

            # Duplicate emp_code check inside uploaded file
            cleaned_emp_codes = df["emp_code"].apply(lambda x: clean_text(x))
            duplicate_emp_codes = cleaned_emp_codes[
                cleaned_emp_codes.duplicated() & (cleaned_emp_codes != "")
            ].unique()

            for emp_code in duplicate_emp_codes:
                row_errors.append(f"Duplicate emp_code found in file: {emp_code}")

            valid_genders = ["male", "female", "other"]
            valid_tax_regimes = ["old", "new"]

            for index, row in df.iterrows():
                row_no = index + 2

                emp_code = clean_text(row.get("emp_code"))
                employee_name = clean_text(row.get("employee_name"))
                role = clean_text(row.get("role"))
                gender = clean_text(row.get("gender"), "male").lower()
                tax_regime = clean_text(row.get("tax_regime"), "new").lower()
                monthly_salary = clean_float(row.get("monthly_salary"), 0)
                special_allowance = clean_float(row.get("special_allowance"), 0)
                other_annual_deductions = clean_float(row.get("other_annual_deductions"), 0)

                if emp_code == "":
                    row_errors.append(f"Row {row_no}: emp_code missing")

                if employee_name == "":
                    row_errors.append(f"Row {row_no}: employee_name missing")

                if role == "":
                    row_errors.append(f"Row {row_no}: role missing")

                if monthly_salary <= 0:
                    row_errors.append(f"Row {row_no}: monthly_salary must be greater than 0")

                if gender and gender not in valid_genders:
                    row_errors.append(f"Row {row_no}: gender must be male, female, or other")

                if tax_regime and tax_regime not in valid_tax_regimes:
                    row_errors.append(f"Row {row_no}: tax_regime must be old or new")

                if special_allowance < 0:
                    row_errors.append(f"Row {row_no}: special_allowance cannot be negative")

                if other_annual_deductions < 0:
                    row_errors.append(f"Row {row_no}: other_annual_deductions cannot be negative")

            if row_errors:
                session["error_report"] = create_error_report(
                    row_errors,
                    "employee_upload_errors.xlsx"
                )
                flash("Upload failed. Please download the error report and fix the file.", "danger")
                return redirect(url_for("upload_employees"))

            conn = get_db()
            cur = conn.cursor()

            # Free plan upload limit check - only NEW employees count honge.
            if not is_admin_user():
                active_plan = get_active_plan()

                if not active_plan.get("is_pro"):
                    cur.execute("""
                        SELECT emp_code
                        FROM employees
                        WHERE company_id = ?
                    """, (company_id,))

                    existing_emp_codes = {
                        clean_text(row["emp_code"]).lower()
                        for row in cur.fetchall()
                    }

                    upload_emp_codes = {
                        clean_text(emp_code).lower()
                        for emp_code in df["emp_code"].tolist()
                        if clean_text(emp_code) != ""
                    }

                    new_emp_codes = upload_emp_codes - existing_emp_codes

                    existing_count = len(existing_emp_codes)
                    new_upload_count = len(new_emp_codes)

                    if existing_count + new_upload_count > FREE_EMPLOYEE_LIMIT:
                        flash(
                            f"Free plan allows up to {FREE_EMPLOYEE_LIMIT} employees only. "
                            f"You already have {existing_count} employee(s), and this file will add "
                            f"{new_upload_count} new employee(s). Please upgrade to PRO for unlimited employees.",
                            "warning"
                        )
                        return redirect(url_for("pricing"))

            # Get leave policy for default balances
            cur.execute("""
                SELECT 
                    casual_leave_limit,
                    sick_leave_limit,
                    paid_leave_limit
                FROM leave_policy
                WHERE company_id = ?
                LIMIT 1
            """, (company_id,))

            leave_policy = cur.fetchone()

            if leave_policy:
                default_casual_leave = float(leave_policy["casual_leave_limit"] or 6)
                default_sick_leave = float(leave_policy["sick_leave_limit"] or 6)
                default_paid_leave = float(leave_policy["paid_leave_limit"] or 12)
            else:
                default_casual_leave = 6
                default_sick_leave = 6
                default_paid_leave = 12

            added_count = 0
            updated_count = 0

            for _, row in df.iterrows():
                emp_code = clean_text(row.get("emp_code"))
                employee_name = clean_text(row.get("employee_name"))
                role = clean_text(row.get("role"))
                department = clean_text(row.get("department"), "General")
                gender = clean_text(row.get("gender"), "male").lower()
                monthly_salary = clean_float(row.get("monthly_salary"), 0)
                tax_regime = clean_text(row.get("tax_regime"), "new").lower()
                other_annual_deductions = clean_float(row.get("other_annual_deductions"), 0)
                special_allowance = clean_float(row.get("special_allowance"), 0)

                uan_no = clean_text(row.get("uan_no"))
                esic_no = clean_text(row.get("esic_no"))
                bank_name = clean_text(row.get("bank_name"))
                account_no = clean_text(row.get("account_no"))
                ifsc_code = clean_text(row.get("ifsc_code")).upper()

                if gender not in ["male", "female", "other"]:
                    gender = "male"

                if tax_regime not in ["old", "new"]:
                    tax_regime = "new"

                if department == "":
                    department = "General"

                cur.execute("""
                    SELECT id
                    FROM employees
                    WHERE company_id = ?
                      AND emp_code = ?
                """, (company_id, emp_code))

                existing_employee = cur.fetchone()

                if existing_employee:
                    updated_count += 1
                else:
                    added_count += 1

                cur.execute("""
                    INSERT OR REPLACE INTO employees
                    (
                        company_id,
                        emp_code,
                        employee_name,
                        role,
                        department,
                        gender,
                        monthly_salary,
                        tax_regime,
                        other_annual_deductions,
                        special_allowance,
                        uan_no,
                        esic_no,
                        bank_name,
                        account_no,
                        ifsc_code
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    company_id,
                    emp_code,
                    employee_name,
                    role,
                    department,
                    gender,
                    monthly_salary,
                    tax_regime,
                    other_annual_deductions,
                    special_allowance,
                    uan_no,
                    esic_no,
                    bank_name,
                    account_no,
                    ifsc_code
                ))

                # Create leave balance only for new employees / missing balance
                cur.execute("""
                    SELECT id
                    FROM leave_balances
                    WHERE company_id = ?
                      AND emp_code = ?
                """, (company_id, emp_code))

                existing_balance = cur.fetchone()

                if not existing_balance:
                    cur.execute("""
                        INSERT INTO leave_balances
                        (
                            company_id,
                            emp_code,
                            casual_leave,
                            sick_leave,
                            paid_leave,
                            used_leave
                        )
                        VALUES (?, ?, ?, ?, ?, 0)
                    """, (
                        company_id,
                        emp_code,
                        default_casual_leave,
                        default_sick_leave,
                        default_paid_leave
                    ))

            conn.commit()
            session.pop("error_report", None)

            flash(
                f"Employee master uploaded successfully. Added: {added_count}, Updated: {updated_count}. Leave balances checked/created.",
                "success"
            )

            return redirect(url_for("employees_list"))

        except Exception as e:
            if conn:
                conn.rollback()

            flash(f"Upload failed: {str(e)}", "danger")
            return redirect(url_for("upload_employees"))

        finally:
            if conn:
                conn.close()

    return render_template("upload_employees.html")


@app.route("/employees")
@login_required
def employees_list():
    department = request.args.get("department", "").strip()
    search = request.args.get("search", "").strip()

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    query = """
        SELECT *
        FROM employees
        WHERE company_id = ?
    """

    params = [company_id]

    if department:
        query += " AND department = ?"
        params.append(department)

    if search:
        query += """
            AND (
                emp_code LIKE ?
                OR employee_name LIKE ?
                OR role LIKE ?
                OR department LIKE ?
            )
        """
        search_value = f"%{search}%"
        params.extend([search_value, search_value, search_value, search_value])

    query += " ORDER BY id DESC"

    cur.execute(query, tuple(params))
    employees = cur.fetchall()

    cur.execute("""
        SELECT DISTINCT department
        FROM employees
        WHERE company_id = ?
          AND department IS NOT NULL
          AND department != ''
        ORDER BY department
    """, (company_id,))
    departments = cur.fetchall()

    cur.execute("""
        SELECT COUNT(*) AS total_employees
        FROM employees
        WHERE company_id = ?
    """, (company_id,))
    total_employees = cur.fetchone()["total_employees"] or 0

    cur.execute("""
        SELECT COUNT(DISTINCT department) AS total_departments
        FROM employees
        WHERE company_id = ?
          AND department IS NOT NULL
          AND department != ''
    """, (company_id,))
    total_departments = cur.fetchone()["total_departments"] or 0

    cur.execute("""
        SELECT COALESCE(SUM(monthly_salary), 0) AS total_salary
        FROM employees
        WHERE company_id = ?
    """, (company_id,))
    total_salary = round(float(cur.fetchone()["total_salary"] or 0))

    cur.execute("""
        SELECT COUNT(*) AS male_count
        FROM employees
        WHERE company_id = ?
          AND LOWER(COALESCE(gender, '')) = 'male'
    """, (company_id,))
    male_count = cur.fetchone()["male_count"] or 0

    cur.execute("""
        SELECT COUNT(*) AS female_count
        FROM employees
        WHERE company_id = ?
          AND LOWER(COALESCE(gender, '')) = 'female'
    """, (company_id,))
    female_count = cur.fetchone()["female_count"] or 0

    conn.close()

    return render_template(
        "employees.html",
        employees=employees,
        departments=departments,

        selected_department=department,
        search=search,

        total_employees=total_employees,
        total_departments=total_departments,
        total_salary=total_salary,
        male_count=male_count,
        female_count=female_count
    )

@app.route("/download-employee-sample")
@login_required
def download_employee_sample():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    data = {
        "emp_code": ["EMP001", "EMP002"],
        "employee_name": ["Rahul Sharma", "Priya Singh"],
        "role": ["Manager", "HR Executive"],
        "department": ["Operations", "HR"],
        "gender": ["male", "female"],
        "monthly_salary": [25000, 30000],
        "tax_regime": ["new", "new"],
        "other_annual_deductions": [0, 0],
        "special_allowance": [2000, 3000],
        "uan_no": ["123456789012", "222233334444"],
        "esic_no": ["9876543210", "111122223333"],
        "bank_name": ["HDFC Bank", "ICICI Bank"],
        "account_no": ["50100234567890", "123456789012"],
        "ifsc_code": ["HDFC0001234", "ICIC0005678"]
    }

    df = pd.DataFrame(data)

    file_path = os.path.join(UPLOAD_FOLDER, "employee_master_sample.xlsx")

    instructions = pd.DataFrame({
        "Field": [
            "emp_code",
            "employee_name",
            "role",
            "department",
            "gender",
            "monthly_salary",
            "tax_regime",
            "other_annual_deductions",
            "special_allowance",
            "uan_no",
            "esic_no",
            "bank_name",
            "account_no",
            "ifsc_code"
        ],
        "Required": [
            "Yes",
            "Yes",
            "Yes",
            "Optional",
            "Optional",
            "Yes",
            "Optional",
            "Optional",
            "Optional",
            "Optional",
            "Optional",
            "Optional",
            "Optional",
            "Optional"
        ],
        "Example": [
            "EMP001",
            "Rahul Sharma",
            "Manager",
            "Operations",
            "male / female / other",
            "25000",
            "old / new",
            "0",
            "2000",
            "123456789012",
            "9876543210",
            "HDFC Bank",
            "50100234567890",
            "HDFC0001234"
        ],
        "Notes": [
            "Must be unique for each employee. Existing emp_code will update employee details.",
            "Employee full name.",
            "Designation or job role.",
            "Blank department will be treated as General.",
            "Use male, female, or other. Blank value defaults to male.",
            "Monthly salary must be greater than 0.",
            "Use old or new. Blank value defaults to new.",
            "Use 0 if not applicable.",
            "Use 0 if not applicable.",
            "Keep as text to avoid number formatting issues.",
            "Keep as text to avoid number formatting issues.",
            "Employee bank name.",
            "Keep as text to avoid number formatting issues.",
            "Use uppercase IFSC code."
        ]
    })

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Employee Master")
        instructions.to_excel(writer, index=False, sheet_name="Instructions")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        required_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        optional_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0")
        )

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for row_cells in ws.iter_rows(min_row=2):
                for cell in row_cells:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")

            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(cell_value))

                ws.column_dimensions[column_letter].width = max_length + 4

        # Keep emp_code, UAN, ESIC, account no and IFSC as text in Employee Master sheet
        employee_ws = workbook["Employee Master"]
        text_columns = ["A", "J", "K", "M", "N"]

        for col in text_columns:
            for cell in employee_ws[col]:
                cell.number_format = "@"

        # Format instruction required/optional rows
        instruction_ws = workbook["Instructions"]

        for row_idx in range(2, instruction_ws.max_row + 1):
            required_value = str(instruction_ws.cell(row=row_idx, column=2).value or "").strip().lower()

            if required_value == "yes":
                fill = required_fill
            else:
                fill = optional_fill

            for col_idx in range(1, instruction_ws.max_column + 1):
                instruction_ws.cell(row=row_idx, column=col_idx).fill = fill

    return send_file(
        file_path,
        as_attachment=True,
        download_name="employee_master_sample.xlsx"
    )


# ---------------------------
# ATTENDANCE
# ---------------------------
@app.route("/upload-attendance", methods=["GET", "POST"])
@login_required
def upload_attendance():
    if request.method == "POST":
        company_id = current_company_id()

        if not company_id:
            flash("Company not found. Please login again.", "danger")
            return redirect(url_for("login"))

        if "file" not in request.files:
            flash("Please select a file.", "warning")
            return redirect(url_for("upload_attendance"))

        file = request.files["file"]

        if not file or file.filename.strip() == "":
            flash("Please select a file.", "warning")
            return redirect(url_for("upload_attendance"))

        filename = file.filename.lower()

        if not (filename.endswith(".csv") or filename.endswith(".xlsx")):
            flash("Only CSV or Excel (.xlsx) file allowed.", "danger")
            return redirect(url_for("upload_attendance"))

        path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
        file.save(path)

        conn = None

        try:
            if filename.endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path, engine="openpyxl")

            # Clean column names
            df.columns = [str(col).strip().lower() for col in df.columns]

            required_columns = [
                "emp_code",
                "month",
                "working_days",
                "present_days",
                "weekly_off",
                "paid_leave",
                "holiday",
                "lop_days",
                "overtime_hours"
            ]

            # Required column validation - helper function dependency removed
            missing_columns = [
                col for col in required_columns
                if col not in df.columns
            ]

            if missing_columns:
                session["error_report"] = create_error_report(
                    [f"Missing required column: {col}" for col in missing_columns],
                    "attendance_upload_errors.xlsx"
                )
                flash("Upload failed. Required columns are missing. Please download the error report.", "danger")
                return redirect(url_for("upload_attendance"))

            row_errors = []

            df["emp_code_clean"] = df["emp_code"].apply(lambda x: clean_text(x))
            df["month_clean"] = df["month"].apply(lambda x: clean_text(x))

            duplicate_rows = df[
                df.duplicated(subset=["emp_code_clean", "month_clean"], keep=False)
                & (df["emp_code_clean"] != "")
                & (df["month_clean"] != "")
            ]

            if not duplicate_rows.empty:
                duplicate_pairs = duplicate_rows[["emp_code_clean", "month_clean"]].drop_duplicates()

                for _, dup in duplicate_pairs.iterrows():
                    row_errors.append(
                        f"Duplicate attendance found for emp_code {dup['emp_code_clean']} in month {dup['month_clean']}"
                    )

            for index, row in df.iterrows():
                row_no = index + 2

                emp_code = clean_text(row.get("emp_code"))
                month = clean_text(row.get("month"))

                working_days = clean_float(row.get("working_days"), -1)
                present_days = clean_float(row.get("present_days"), -1)
                weekly_off = clean_float(row.get("weekly_off"), -1)
                paid_leave = clean_float(row.get("paid_leave"), -1)
                holiday = clean_float(row.get("holiday"), -1)
                lop_days = clean_float(row.get("lop_days"), -1)
                paid_days = clean_float(row.get("paid_days"), 0)
                overtime_hours = clean_float(row.get("overtime_hours"), -1)
                bonus = clean_float(row.get("bonus"), 0)
                manual_deduction = clean_float(row.get("manual_deduction"), 0)

                if emp_code == "":
                    row_errors.append(f"Row {row_no}: emp_code missing")

                if month == "":
                    row_errors.append(f"Row {row_no}: month missing")
                else:
                    try:
                        datetime.datetime.strptime(month, "%Y-%m")
                    except Exception:
                        row_errors.append(f"Row {row_no}: month must be in YYYY-MM format, example 2026-12")

                if working_days <= 0:
                    row_errors.append(f"Row {row_no}: working_days must be greater than 0")

                if present_days < 0:
                    row_errors.append(f"Row {row_no}: present_days cannot be negative")

                if weekly_off < 0:
                    row_errors.append(f"Row {row_no}: weekly_off cannot be negative")

                if paid_leave < 0:
                    row_errors.append(f"Row {row_no}: paid_leave cannot be negative")

                if holiday < 0:
                    row_errors.append(f"Row {row_no}: holiday cannot be negative")

                if lop_days < 0:
                    row_errors.append(f"Row {row_no}: lop_days cannot be negative")

                if paid_days < 0:
                    row_errors.append(f"Row {row_no}: paid_days cannot be negative")

                if overtime_hours < 0:
                    row_errors.append(f"Row {row_no}: overtime_hours cannot be negative")

                if bonus < 0:
                    row_errors.append(f"Row {row_no}: bonus cannot be negative")

                if manual_deduction < 0:
                    row_errors.append(f"Row {row_no}: manual_deduction cannot be negative")

                if working_days > 31:
                    row_errors.append(f"Row {row_no}: working_days cannot be greater than 31")

                if present_days > 31:
                    row_errors.append(f"Row {row_no}: present_days cannot be greater than 31")

                if weekly_off > 31:
                    row_errors.append(f"Row {row_no}: weekly_off cannot be greater than 31")

                if paid_leave > 31:
                    row_errors.append(f"Row {row_no}: paid_leave cannot be greater than 31")

                if holiday > 31:
                    row_errors.append(f"Row {row_no}: holiday cannot be greater than 31")

                if lop_days > 31:
                    row_errors.append(f"Row {row_no}: lop_days cannot be greater than 31")

                calculated_paid_days = present_days + weekly_off + paid_leave + holiday - lop_days

                if calculated_paid_days < 0:
                    row_errors.append(f"Row {row_no}: calculated paid_days cannot be negative")

                if calculated_paid_days > 31:
                    row_errors.append(f"Row {row_no}: calculated paid_days cannot be greater than 31")

                if paid_days > 0 and abs(paid_days - calculated_paid_days) > 0.01:
                    row_errors.append(
                        f"Row {row_no}: paid_days mismatch. Expected {calculated_paid_days}, found {paid_days}"
                    )

            if row_errors:
                session["error_report"] = create_error_report(
                    row_errors,
                    "attendance_upload_errors.xlsx"
                )
                flash("Upload failed. Please download the error report and fix the file.", "danger")
                return redirect(url_for("upload_attendance"))

            conn = get_db()
            cur = conn.cursor()

            missing_employee_errors = []

            for index, row in df.iterrows():
                row_no = index + 2
                emp_code = clean_text(row.get("emp_code"))

                cur.execute("""
                    SELECT id
                    FROM employees
                    WHERE company_id = ?
                      AND emp_code = ?
                """, (company_id, emp_code))

                if not cur.fetchone():
                    missing_employee_errors.append(
                        f"Row {row_no}: Employee code not found in Employee Master: {emp_code}"
                    )

            if missing_employee_errors:
                session["error_report"] = create_error_report(
                    missing_employee_errors,
                    "attendance_upload_errors.xlsx"
                )
                flash("Upload failed. Some employee codes were not found. Please download the error report.", "danger")
                return redirect(url_for("upload_attendance"))

            uploaded_months = df["month_clean"].dropna().unique().tolist()

            for uploaded_month in uploaded_months:
                cur.execute("""
                    DELETE FROM attendance
                    WHERE company_id = ?
                      AND month = ?
                """, (company_id, uploaded_month))

            success_count = 0

            for _, row in df.iterrows():
                emp_code = clean_text(row.get("emp_code"))
                month = clean_text(row.get("month"))

                working_days = clean_float(row.get("working_days"), 0)
                present_days = clean_float(row.get("present_days"), 0)
                weekly_off = clean_float(row.get("weekly_off"), 0)
                paid_leave = clean_float(row.get("paid_leave"), 0)
                holiday = clean_float(row.get("holiday"), 0)
                lop_days = clean_float(row.get("lop_days"), 0)
                paid_days = clean_float(row.get("paid_days"), 0)

                if paid_days <= 0:
                    paid_days = present_days + weekly_off + paid_leave + holiday - lop_days

                overtime_hours = clean_float(row.get("overtime_hours"), 0)
                bonus = clean_float(row.get("bonus"), 0)
                manual_deduction = clean_float(row.get("manual_deduction"), 0)

                cur.execute("""
                    INSERT INTO attendance
                    (
                        company_id,
                        emp_code,
                        month,
                        working_days,
                        present_days,
                        weekly_off,
                        paid_leave,
                        holiday,
                        lop_days,
                        paid_days,
                        overtime_hours,
                        bonus,
                        manual_deduction
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    company_id,
                    emp_code,
                    month,
                    working_days,
                    present_days,
                    weekly_off,
                    paid_leave,
                    holiday,
                    lop_days,
                    paid_days,
                    overtime_hours,
                    bonus,
                    manual_deduction
                ))

                success_count += 1

            conn.commit()
            session.pop("error_report", None)

            flash(
                f"Attendance uploaded successfully. Added: {success_count}. Existing attendance for uploaded month(s) was replaced.",
                "success"
            )

            return redirect(url_for("run_payroll"))

        except Exception as e:
            if conn:
                conn.rollback()

            flash(f"Upload failed: {str(e)}", "danger")
            return redirect(url_for("upload_attendance"))

        finally:
            if conn:
                conn.close()

    return render_template("upload_attendance.html")


@app.route("/attendance")
@login_required
def attendance_list():
    month = request.args.get("month", "").strip()
    department = request.args.get("department", "").strip()
    search = request.args.get("search", "").strip()

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    query = """
        SELECT 
            a.*,
            COALESCE(e.employee_name, '') AS employee_name,
            COALESCE(e.role, '') AS role,
            COALESCE(e.department, '') AS department,
            COALESCE(e.gender, '') AS gender
        FROM attendance a
        LEFT JOIN employees e
          ON a.company_id = e.company_id
         AND a.emp_code = e.emp_code
        WHERE a.company_id = ?
    """

    params = [company_id]

    if month:
        query += " AND a.month = ?"
        params.append(month)

    if department:
        query += " AND e.department = ?"
        params.append(department)

    if search:
        query += """
            AND (
                a.emp_code LIKE ?
                OR e.employee_name LIKE ?
                OR e.role LIKE ?
                OR e.department LIKE ?
            )
        """
        search_value = f"%{search}%"
        params.extend([search_value, search_value, search_value, search_value])

    query += " ORDER BY a.month DESC, a.id DESC"

    cur.execute(query, tuple(params))
    rows = cur.fetchall()

    cur.execute("""
        SELECT DISTINCT department
        FROM employees
        WHERE company_id = ?
          AND department IS NOT NULL
          AND department != ''
        ORDER BY department
    """, (company_id,))
    departments = cur.fetchall()

    cur.execute("""
        SELECT DISTINCT month
        FROM attendance
        WHERE company_id = ?
          AND month IS NOT NULL
          AND month != ''
        ORDER BY month DESC
    """, (company_id,))
    months = cur.fetchall()

    total_records = len(rows)

    total_working_days = round(sum(float(r["working_days"] or 0) for r in rows), 2)
    total_present_days = round(sum(float(r["present_days"] or 0) for r in rows), 2)

    total_weekly_off = round(sum(float(r["weekly_off"] or 0) for r in rows), 2)
    total_paid_leave = round(sum(float(r["paid_leave"] or 0) for r in rows), 2)
    total_holiday = round(sum(float(r["holiday"] or 0) for r in rows), 2)
    total_lop_days = round(sum(float(r["lop_days"] or 0) for r in rows), 2)
    total_paid_days = round(sum(float(r["paid_days"] or 0) for r in rows), 2)

    total_overtime_hours = round(sum(float(r["overtime_hours"] or 0) for r in rows), 2)
    total_bonus = round(sum(float(r["bonus"] or 0) for r in rows))
    total_manual_deduction = round(sum(float(r["manual_deduction"] or 0) for r in rows))

    total_absent_days = round(
        total_working_days
        - total_present_days
        - total_weekly_off
        - total_paid_leave
        - total_holiday,
        2
    )

    if total_absent_days < 0:
        total_absent_days = 0

    if total_working_days > 0:
        attendance_percentage = round((total_present_days / total_working_days) * 100, 2)
        paid_days_percentage = round((total_paid_days / total_working_days) * 100, 2)
    else:
        attendance_percentage = 0
        paid_days_percentage = 0

    conn.close()

    return render_template(
        "attendance.html",
        rows=rows,
        departments=departments,
        months=months,

        selected_month=month,
        selected_department=department,
        search=search,

        total_records=total_records,
        total_working_days=total_working_days,
        total_present_days=total_present_days,
        total_weekly_off=total_weekly_off,
        total_paid_leave=total_paid_leave,
        total_holiday=total_holiday,
        total_lop_days=total_lop_days,
        total_paid_days=total_paid_days,
        total_absent_days=total_absent_days,
        total_overtime_hours=total_overtime_hours,
        total_bonus=total_bonus,
        total_manual_deduction=total_manual_deduction,
        attendance_percentage=attendance_percentage,
        paid_days_percentage=paid_days_percentage
    )


@app.route("/leave-management", methods=["GET", "POST"])
@login_required
def leave_management():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    try:
        # Ensure default leave policy exists
        cur.execute("""
            INSERT OR IGNORE INTO leave_policy_settings
            (
                company_id,
                casual_leave_limit,
                sick_leave_limit,
                paid_leave_limit
            )
            VALUES (?, 6, 6, 12)
        """, (company_id,))

        conn.commit()

        cur.execute("""
            SELECT 
                casual_leave_limit,
                sick_leave_limit,
                paid_leave_limit
            FROM leave_policy_settings
            WHERE company_id = ?
        """, (company_id,))

        leave_policy = cur.fetchone()

        if request.method == "POST":
            emp_code = request.form.get("emp_code", "").strip()
            leave_type = request.form.get("leave_type", "").strip()
            start_date = request.form.get("start_date", "").strip()
            end_date = request.form.get("end_date", "").strip()
            total_days_input = request.form.get("total_days", "").strip()
            reason = request.form.get("reason", "").strip()

            errors = []

            if not emp_code:
                errors.append("Employee is required.")

            if not leave_type:
                errors.append("Leave type is required.")

            allowed_leave_types = [
                "Casual Leave",
                "Sick Leave",
                "Paid Leave",
                "Leave Without Pay",
                "Unpaid Leave",
                "LWP"
            ]

            if leave_type and leave_type not in allowed_leave_types:
                errors.append("Invalid leave type selected.")

            start_date_obj = None
            end_date_obj = None

            if not start_date:
                errors.append("Start date is required.")
            else:
                try:
                    start_date_obj = datetime.datetime.strptime(start_date, "%Y-%m-%d")
                except Exception:
                    errors.append("Start date must be a valid date.")

            if not end_date:
                errors.append("End date is required.")
            else:
                try:
                    end_date_obj = datetime.datetime.strptime(end_date, "%Y-%m-%d")
                except Exception:
                    errors.append("End date must be a valid date.")

            if start_date_obj and end_date_obj and end_date_obj < start_date_obj:
                errors.append("End date cannot be earlier than start date.")

            # Auto calculate total days if blank or 0
            total_days = 0

            try:
                total_days = float(total_days_input or 0)
            except Exception:
                total_days = 0

            if total_days <= 0 and start_date_obj and end_date_obj:
                total_days = (end_date_obj - start_date_obj).days + 1

            if total_days <= 0:
                errors.append("Total leave days must be greater than 0.")

            if total_days > 365:
                errors.append("Total leave days cannot be greater than 365.")

            # Employee exists check
            if emp_code:
                cur.execute("""
                    SELECT emp_code
                    FROM employees
                    WHERE company_id = ?
                      AND emp_code = ?
                """, (company_id, emp_code))

                emp = cur.fetchone()

                if not emp:
                    errors.append("Employee code not found in Employee Master.")

            if errors:
                flash(" ".join(errors), "danger")
                return redirect(url_for("leave_management"))

            # Duplicate / overlapping leave request check
            # Same employee ke liye same date range me Pending ya Approved leave dobara create nahi hoga.
            cur.execute("""
                SELECT id, leave_type, start_date, end_date, status
                FROM leave_requests
                WHERE company_id = ?
                  AND emp_code = ?
                  AND status IN ('Pending', 'Approved')
                  AND (
                    date(start_date) <= date(?)
                  AND date(end_date) >= date(?)
                  )
                LIMIT 1
            """, (
    company_id,
    emp_code,
    end_date,
    start_date
))

            existing_leave = cur.fetchone()

            if existing_leave:
                flash(
                    f"Leave request already exists for this employee between "
                    f"{existing_leave['start_date']} and {existing_leave['end_date']} "
                    f"with status {existing_leave['status']}. Duplicate leave not allowed.",
                    "warning"
                )
                return redirect(url_for("leave_management"))

            cur.execute("""
                INSERT INTO leave_requests
                (
                    company_id,
                    emp_code,
                    leave_type,
                    start_date,
                    end_date,
                    total_days,
                    reason,
                    status
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 'Pending')
            """, (
                company_id,
                emp_code,
                leave_type,
                start_date,
                end_date,
                total_days,
                reason
            ))

            conn.commit()

            flash("Leave request added successfully.", "success")
            return redirect(url_for("leave_management"))

        # Employees list
        cur.execute("""
            SELECT emp_code, employee_name, department
            FROM employees
            WHERE company_id = ?
            ORDER BY employee_name
        """, (company_id,))

        employees = cur.fetchall()

        # Create missing leave balances only
        for emp in employees:
            cur.execute("""
                INSERT OR IGNORE INTO leave_balances
                (
                    company_id,
                    emp_code,
                    casual_leave,
                    sick_leave,
                    paid_leave,
                    used_leave
                )
                VALUES (?, ?, ?, ?, ?, 0)
            """, (
                company_id,
                emp["emp_code"],
                float(leave_policy["casual_leave_limit"] or 6),
                float(leave_policy["sick_leave_limit"] or 6),
                float(leave_policy["paid_leave_limit"] or 12)
            ))

        conn.commit()

        cur.execute("""
            SELECT 
                lr.*,
                e.employee_name,
                e.department
            FROM leave_requests lr
            LEFT JOIN employees e
              ON lr.company_id = e.company_id
             AND lr.emp_code = e.emp_code
            WHERE lr.company_id = ?
            ORDER BY lr.id DESC
        """, (company_id,))

        leave_requests = cur.fetchall()

        cur.execute("""
            SELECT 
                lb.*,
                e.employee_name,
                e.department
            FROM leave_balances lb
            LEFT JOIN employees e
              ON lb.company_id = e.company_id
             AND lb.emp_code = e.emp_code
            WHERE lb.company_id = ?
            ORDER BY e.employee_name
        """, (company_id,))

        leave_balances = cur.fetchall()

        return render_template(
            "leave_management.html",
            employees=employees,
            leave_requests=leave_requests,
            leave_balances=leave_balances,
            leave_policy=leave_policy
        )

    except Exception as e:
        conn.rollback()
        flash(f"Error in Leave Management: {str(e)}", "danger")
        return redirect(url_for("dashboard"))

    finally:
        conn.close()


@app.route("/update-leave-policy", methods=["POST"])
@login_required
def update_leave_policy():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    def to_float(value, default=0):
        try:
            if value is None or str(value).strip() == "":
                return default
            return float(value)
        except Exception:
            return default

    casual_leave_limit = to_float(request.form.get("casual_leave_limit"), 0)
    sick_leave_limit = to_float(request.form.get("sick_leave_limit"), 0)
    paid_leave_limit = to_float(request.form.get("paid_leave_limit"), 0)

    errors = []

    if casual_leave_limit < 0:
        errors.append("Casual Leave limit cannot be negative.")

    if sick_leave_limit < 0:
        errors.append("Sick Leave limit cannot be negative.")

    if paid_leave_limit < 0:
        errors.append("Paid Leave limit cannot be negative.")

    if casual_leave_limit > 365:
        errors.append("Casual Leave limit cannot be greater than 365.")

    if sick_leave_limit > 365:
        errors.append("Sick Leave limit cannot be greater than 365.")

    if paid_leave_limit > 365:
        errors.append("Paid Leave limit cannot be greater than 365.")

    if errors:
        flash(" ".join(errors), "danger")
        return redirect(url_for("leave_management"))

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO leave_policy_settings
            (
                company_id,
                casual_leave_limit,
                sick_leave_limit,
                paid_leave_limit,
                updated_at
            )
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(company_id) DO UPDATE SET
                casual_leave_limit = excluded.casual_leave_limit,
                sick_leave_limit = excluded.sick_leave_limit,
                paid_leave_limit = excluded.paid_leave_limit,
                updated_at = CURRENT_TIMESTAMP
        """, (
            company_id,
            casual_leave_limit,
            sick_leave_limit,
            paid_leave_limit
        ))

        # Create leave balances for employees who do not have balance records yet.
        # Existing balances are not overwritten.
        cur.execute("""
            SELECT emp_code
            FROM employees
            WHERE company_id = ?
        """, (company_id,))

        employees = cur.fetchall()

        for emp in employees:
            cur.execute("""
                INSERT OR IGNORE INTO leave_balances
                (
                    company_id,
                    emp_code,
                    casual_leave,
                    sick_leave,
                    paid_leave,
                    used_leave
                )
                VALUES (?, ?, ?, ?, ?, 0)
            """, (
                company_id,
                emp["emp_code"],
                casual_leave_limit,
                sick_leave_limit,
                paid_leave_limit
            ))

        conn.commit()

        flash("Leave policy updated successfully. Existing leave balances were not overwritten.", "success")

    except Exception as e:
        conn.rollback()
        flash(f"Error while updating leave policy: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("leave_management"))


@app.route("/approve-leave/<int:leave_id>", methods=["POST"])
@login_required
def approve_leave(leave_id):
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT *
            FROM leave_requests
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        leave = cur.fetchone()

        if not leave:
            flash("Leave request not found.", "warning")
            return redirect(url_for("leave_management"))

        current_status = str(leave["status"] or "").strip()

        if current_status == "Approved":
            flash("This leave is already approved.", "warning")
            return redirect(url_for("leave_management"))

        if current_status == "Rejected":
            flash("Rejected leave cannot be approved directly. Please create a new leave request.", "warning")
            return redirect(url_for("leave_management"))

        emp_code = str(leave["emp_code"] or "").strip()
        leave_type = str(leave["leave_type"] or "").strip()
        total_days = float(leave["total_days"] or 0)

        if not emp_code:
            flash("Employee code missing in leave request.", "danger")
            return redirect(url_for("leave_management"))

        if total_days <= 0:
            flash("Leave days must be greater than 0.", "danger")
            return redirect(url_for("leave_management"))

        # Paid leave types reduce balance.
        leave_column_map = {
            "Casual Leave": "casual_leave",
            "Sick Leave": "sick_leave",
            "Paid Leave": "paid_leave"
        }

        # LWP / unpaid leave does not reduce leave balance.
        unpaid_leave_types = [
            "Leave Without Pay",
            "Unpaid Leave",
            "LWP"
        ]

        if leave_type in unpaid_leave_types:
            cur.execute("""
                UPDATE leave_requests
                SET status = 'Approved'
                WHERE id = ?
                  AND company_id = ?
            """, (leave_id, company_id))

            conn.commit()

            flash("Leave Without Pay approved successfully. Leave balance was not changed.", "success")
            return redirect(url_for("leave_management"))

        if leave_type not in leave_column_map:
            flash(f"Invalid leave type: {leave_type}", "danger")
            return redirect(url_for("leave_management"))

        balance_column = leave_column_map[leave_type]

        cur.execute("""
            SELECT *
            FROM leave_balances
            WHERE company_id = ?
              AND emp_code = ?
        """, (company_id, emp_code))

        balance_row = cur.fetchone()

        if not balance_row:
            flash("Leave balance not found for this employee. Please check leave balance setup.", "danger")
            return redirect(url_for("leave_management"))

        current_balance = float(balance_row[balance_column] or 0)
        used_leave = float(balance_row["used_leave"] or 0)

        if current_balance < total_days:
            flash(
                f"Insufficient {leave_type} balance. Available: {current_balance}, Required: {total_days}.",
                "danger"
            )
            return redirect(url_for("leave_management"))

        new_balance = current_balance - total_days
        new_used_leave = used_leave + total_days

        cur.execute(f"""
            UPDATE leave_balances
            SET {balance_column} = ?,
                used_leave = ?
            WHERE company_id = ?
              AND emp_code = ?
        """, (new_balance, new_used_leave, company_id, emp_code))

        cur.execute("""
            UPDATE leave_requests
            SET status = 'Approved'
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        conn.commit()

        flash(
            f"Leave approved successfully. {leave_type} balance updated: {current_balance} → {new_balance}.",
            "success"
        )

    except Exception as e:
        conn.rollback()
        flash(f"Error while approving leave: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("leave_management"))


@app.route("/reject-leave/<int:leave_id>", methods=["POST"])
@login_required
def reject_leave(leave_id):
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT *
            FROM leave_requests
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        leave = cur.fetchone()

        if not leave:
            flash("Leave request not found.", "warning")
            return redirect(url_for("leave_management"))

        current_status = str(leave["status"] or "").strip()

        if current_status == "Rejected":
            flash("This leave is already rejected.", "warning")
            return redirect(url_for("leave_management"))

        if current_status == "Approved":
            flash("Approved leave cannot be rejected directly because leave balance is already updated. Delete/cancel reversal logic is required.", "warning")
            return redirect(url_for("leave_management"))

        cur.execute("""
            UPDATE leave_requests
            SET status = 'Rejected'
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        conn.commit()

        flash("Leave rejected successfully. Leave balance was not changed.", "success")

    except Exception as e:
        conn.rollback()
        flash(f"Error while rejecting leave: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("leave_management"))


@app.route("/cancel-leave/<int:leave_id>", methods=["POST"])
@login_required
def cancel_leave(leave_id):
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT *
            FROM leave_requests
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        leave = cur.fetchone()

        if not leave:
            flash("Leave request not found.", "warning")
            return redirect(url_for("leave_management"))

        emp_code = str(leave["emp_code"] or "").strip()
        leave_type = str(leave["leave_type"] or "").strip()
        status = str(leave["status"] or "").strip()
        total_days = float(leave["total_days"] or 0)

        leave_column_map = {
            "Casual Leave": "casual_leave",
            "Sick Leave": "sick_leave",
            "Paid Leave": "paid_leave"
        }

        unpaid_leave_types = [
            "Leave Without Pay",
            "Unpaid Leave",
            "LWP"
        ]

        # Approved paid leave cancel/reverse karega
        if status == "Approved" and leave_type in leave_column_map:
            balance_column = leave_column_map[leave_type]

            cur.execute("""
                SELECT *
                FROM leave_balances
                WHERE company_id = ?
                  AND emp_code = ?
            """, (company_id, emp_code))

            balance_row = cur.fetchone()

            if not balance_row:
                flash("Leave balance not found. Cannot reverse approved leave.", "danger")
                return redirect(url_for("leave_management"))

            current_balance = float(balance_row[balance_column] or 0)
            used_leave = float(balance_row["used_leave"] or 0)

            new_balance = current_balance + total_days
            new_used_leave = used_leave - total_days

            if new_used_leave < 0:
                new_used_leave = 0

            cur.execute(f"""
                UPDATE leave_balances
                SET {balance_column} = ?,
                    used_leave = ?
                WHERE company_id = ?
                  AND emp_code = ?
            """, (
                new_balance,
                new_used_leave,
                company_id,
                emp_code
            ))

        # Approved LWP / Pending / Rejected delete only, balance unchanged
        elif status == "Approved" and leave_type in unpaid_leave_types:
            pass

        elif status in ["Pending", "Rejected"]:
            pass

        else:
            flash("Invalid leave status. Cannot cancel this leave request.", "warning")
            return redirect(url_for("leave_management"))

        cur.execute("""
            DELETE FROM leave_requests
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        conn.commit()

        flash("Leave request cancelled/deleted successfully.", "success")

    except Exception as e:
        conn.rollback()
        flash(f"Error while cancelling leave: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("leave_management"))


@app.route("/download-attendance-sample")
@login_required
def download_attendance_sample():
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    data = {
        "emp_code": ["EMP001", "EMP002"],
        "month": ["2026-12", "2026-12"],
        "working_days": [30, 30],
        "present_days": [26, 24],
        "weekly_off": [4, 4],
        "paid_leave": [0, 1],
        "holiday": [0, 0],
        "lop_days": [0, 1],
        "paid_days": [30, 28],
        "overtime_hours": [2, 2],
        "bonus": [0, 0],
        "manual_deduction": [0, 0]
    }

    df = pd.DataFrame(data)

    file_path = os.path.join(UPLOAD_FOLDER, "attendance_sample.xlsx")

    instructions = pd.DataFrame({
        "Field": [
            "emp_code",
            "month",
            "working_days",
            "present_days",
            "weekly_off",
            "paid_leave",
            "holiday",
            "lop_days",
            "paid_days",
            "overtime_hours",
            "bonus",
            "manual_deduction"
        ],
        "Required": [
            "Yes",
            "Yes",
            "Yes",
            "Yes",
            "Yes",
            "Yes",
            "Yes",
            "Yes",
            "Optional",
            "Yes",
            "Optional",
            "Optional"
        ],
        "Example": [
            "EMP001",
            "2026-12",
            "30",
            "26",
            "4",
            "0",
            "0",
            "0",
            "30",
            "2",
            "0",
            "0"
        ],
        "Notes": [
            "Employee code must exist in Employee Master.",
            "Use YYYY-MM format only, example 2026-12.",
            "Month working days / salary days as per company policy.",
            "Actual present days.",
            "Weekly off days. Example: 4 Sundays.",
            "Paid leave days included in salary payable days.",
            "Paid holiday days included in salary payable days.",
            "Loss of Pay / unpaid leave days.",
            "Optional. If blank or 0, system calculates: present_days + weekly_off + paid_leave + holiday - lop_days.",
            "Use 0 if no overtime.",
            "Optional attendance bonus. Use 0 if not applicable.",
            "Optional manual deduction. Use 0 if not applicable."
        ]
    })

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Attendance")
        instructions.to_excel(writer, index=False, sheet_name="Instructions")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        required_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        optional_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0")
        )

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for row_cells in ws.iter_rows(min_row=2):
                for cell in row_cells:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")

            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(cell_value))

                ws.column_dimensions[column_letter].width = max_length + 4

        # Keep emp_code and month as text
        attendance_ws = workbook["Attendance"]

        for col in ["A", "B"]:
            for cell in attendance_ws[col]:
                cell.number_format = "@"

        # Format instruction required/optional rows
        instruction_ws = workbook["Instructions"]

        for row_idx in range(2, instruction_ws.max_row + 1):
            required_value = str(instruction_ws.cell(row=row_idx, column=2).value or "").strip().lower()

            if required_value == "yes":
                fill = required_fill
            else:
                fill = optional_fill

            for col_idx in range(1, instruction_ws.max_column + 1):
                instruction_ws.cell(row=row_idx, column=col_idx).fill = fill

    return send_file(
        file_path,
        as_attachment=True,
        download_name="attendance_sample.xlsx"
    )


@app.route("/download-error-report")
@login_required
def download_error_report():
    file_path = session.get("error_report")
    if not file_path or not os.path.exists(file_path):
        flash("No error report available.")
        return redirect(url_for("dashboard"))
    return send_file(file_path, as_attachment=True)


# ---------------------------
# PAYROLL
# ---------------------------
@app.route("/run-payroll", methods=["GET", "POST"])
@login_required
def run_payroll():
    if request.method == "GET":
        return render_template("run_payroll.html")

    company_id = current_company_id()
    month = request.form.get("month")

    if not month:
        flash("Please select payroll month.")
        return redirect(url_for("run_payroll"))

    conn = get_db()
    cur = conn.cursor()
    run_id = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    # Company overtime policy
    cur.execute("""
        SELECT COALESCE(overtime_multiplier, 1) AS overtime_multiplier
        FROM companies
        WHERE id = ?
    """, (company_id,))
    company = cur.fetchone()

    overtime_multiplier = 1
    if company:
        try:
            overtime_multiplier = float(company["overtime_multiplier"] or 1)
        except Exception:
            overtime_multiplier = 1

    if overtime_multiplier not in [1, 2]:
        overtime_multiplier = 1

    # Compliance settings
    settings = get_compliance_settings(company_id)

    pf_employee_rate = float(settings["pf_employee_rate"] or 12) / 100
    pf_employer_rate = float(settings["pf_employer_rate"] or 12) / 100
    pf_wage_ceiling = float(settings["pf_wage_ceiling"] or 15000)
    pf_max_deduction = float(settings["pf_max_deduction"] or 1800)

    esic_employee_rate = float(settings["esic_employee_rate"] or 0.75) / 100
    esic_employer_rate = float(settings["esic_employer_rate"] or 3.25) / 100
    esic_wage_limit = float(settings["esic_wage_limit"] or 21000)

    gratuity_rate = float(settings["gratuity_rate"] or 4.81) / 100
    bonus_rate = float(settings["bonus_rate"] or 8.33) / 100
    tds_enabled = int(settings["tds_enabled"] or 0)

    salary_days_policy = settings["salary_days_policy"] or "attendance"
    custom_salary_days = float(settings["custom_salary_days"] or 30)

    count_weekly_off_paid = int(settings["count_weekly_off_paid"] or 0)
    count_paid_leave_paid = int(settings["count_paid_leave_paid"] or 0)
    count_holiday_paid = int(settings["count_holiday_paid"] or 0)
    deduct_lop = int(settings["deduct_lop"] or 0)

    festival_bonus_enabled = int(settings["festival_bonus_enabled"] or 0)
    festival_bonus_month = int(settings["festival_bonus_month"] or 10)

    bonus_min_service_days = int(settings["bonus_min_service_days"] or 30)
    bonus_prorata_enabled = int(settings["bonus_prorata_enabled"] or 1)

    # Mark previous payroll of same month as old
    cur.execute("""
        UPDATE payroll_history
        SET is_current = 0
        WHERE company_id = ?
          AND month = ?
    """, (company_id, month))

    # Fetch employees + attendance
    cur.execute("""
        SELECT 
            e.*,
            a.working_days,
            a.present_days,
            COALESCE(a.weekly_off, 0) AS weekly_off,
            COALESCE(a.paid_leave, 0) AS attendance_paid_leave,
            COALESCE(a.holiday, 0) AS holiday,
            COALESCE(a.lop_days, 0) AS attendance_lop_days,
            COALESCE(a.paid_days, 0) AS attendance_paid_days,
            a.overtime_hours,
            a.bonus,
            a.manual_deduction
        FROM employees e
        JOIN attendance a
          ON e.emp_code = a.emp_code
         AND e.company_id = a.company_id
        WHERE e.company_id = ?
          AND a.month = ?
    """, (company_id, month))

    rows = cur.fetchall()

    if not rows:
        conn.close()
        flash("No attendance found for selected month.")
        return redirect(url_for("run_payroll"))

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    year, month_num = map(int, month.split("-"))
    calendar_days = calendar.monthrange(year, month_num)[1]

    # Financial year range
    financial_year_start = f"{year}-04"
    financial_year_end = f"{year + 1}-03"

    if month_num < 4:
        financial_year_start = f"{year - 1}-04"
        financial_year_end = f"{year}-03"

    for row in rows:
        emp_code = row["emp_code"]
        monthly_salary = float(row["monthly_salary"] or 0)

        attendance_working_days = float(row["working_days"] or 30)

        # Salary days policy
        if salary_days_policy == "fixed_26":
            working_days = 26
        elif salary_days_policy == "fixed_30":
            working_days = 30
        elif salary_days_policy == "calendar":
            working_days = calendar_days
        elif salary_days_policy == "custom":
            working_days = custom_salary_days
        else:
            working_days = attendance_working_days

        if working_days <= 0:
            working_days = 30

        present_days = float(row["present_days"] or 0)
        weekly_off = float(row["weekly_off"] or 0)
        attendance_paid_leave = float(row["attendance_paid_leave"] or 0)
        holiday = float(row["holiday"] or 0)
        attendance_lop_days = float(row["attendance_lop_days"] or 0)

        overtime_hours = float(row["overtime_hours"] or 0)
        manual_deduction = float(row["manual_deduction"] or 0)
        attendance_bonus = float(row["bonus"] or 0)

        gender = str(row["gender"] or "male").strip().lower()

        # Approved leave data
        cur.execute("""
            SELECT
                COALESCE(SUM(
                    CASE 
                        WHEN leave_type IN ('Casual Leave', 'Sick Leave', 'Paid Leave')
                        THEN total_days
                        ELSE 0
                    END
                ), 0) AS approved_paid_leave_days,

                COALESCE(SUM(
                    CASE 
                        WHEN leave_type = 'Unpaid Leave'
                        THEN total_days
                        ELSE 0
                    END
                ), 0) AS approved_lwp_days
            FROM leave_requests
            WHERE company_id = ?
              AND emp_code = ?
              AND status = 'Approved'
              AND (
                    substr(start_date, 1, 7) = ?
                    OR substr(end_date, 1, 7) = ?
              )
        """, (company_id, emp_code, month, month))

        leave_data = cur.fetchone()

        approved_paid_leave_days = float(leave_data["approved_paid_leave_days"] or 0)
        approved_lwp_days = float(leave_data["approved_lwp_days"] or 0)

        paid_leave_days = attendance_paid_leave if attendance_paid_leave > 0 else approved_paid_leave_days
        lwp_days = attendance_lop_days if attendance_lop_days > 0 else approved_lwp_days

        # Payable days calculation
        payable_days = present_days

        if count_weekly_off_paid == 1:
            payable_days += weekly_off

        if count_paid_leave_paid == 1:
            payable_days += paid_leave_days

        if count_holiday_paid == 1:
            payable_days += holiday

        if deduct_lop == 1:
            payable_days -= lwp_days

        if payable_days < 0:
            payable_days = 0

        if payable_days > working_days:
            payable_days = working_days

        per_day_salary = monthly_salary / working_days
        lwp_deduction = rupee(per_day_salary * lwp_days)

        earned_salary = per_day_salary * payable_days

        basic = rupee(earned_salary * 0.40)
        da = rupee(earned_salary * 0.10)
        hra = rupee(earned_salary * 0.20)

        special_allowance = float(row["special_allowance"] or 0)
        special_allowance = rupee((special_allowance / working_days) * payable_days)

        other_allowance = earned_salary - basic - da - hra - special_allowance
        if other_allowance < 0:
            other_allowance = 0

        other_allowance = rupee(other_allowance)

        gross = rupee(basic + da + hra + special_allowance + other_allowance)

        # Overtime calculation
        if overtime_hours > 0:
            hourly_rate = monthly_salary / 30 / 8
            overtime_amount = rupee(hourly_rate * overtime_hours * overtime_multiplier)
        else:
            overtime_amount = 0

        # PF calculation with wage ceiling
        pf_base = basic + da

        if pf_wage_ceiling > 0:
            pf_base_for_calculation = min(pf_base, pf_wage_ceiling)
        else:
            pf_base_for_calculation = pf_base

        pf_employee = min(rupee(pf_base_for_calculation * pf_employee_rate), pf_max_deduction)
        pf_employer = min(rupee(pf_base_for_calculation * pf_employer_rate), pf_max_deduction)

        # ESIC calculation
        if gross <= esic_wage_limit:
            esi_employee = round(gross * esic_employee_rate)
            esi_employer = round(gross * esic_employer_rate)
        else:
            esi_employee = 0
            esi_employer = 0

        # Professional tax
        professional_tax = rupee(
            calculate_professional_tax_maharashtra(gross, gender, month)
        )

        # LWF
        lwf = calculate_lwf_maharashtra(month)
        lwf_employee = rupee(lwf["employee"])
        lwf_employer = rupee(lwf["employer"])

        # TDS placeholder
        # Future me actual income tax regime logic add kar sakte hain
        if tds_enabled == 1:
            tds = 0
        else:
            tds = 0

        # Bonus accrual logic
        monthly_bonus_accrual = 0

        if payable_days >= bonus_min_service_days:
            if bonus_prorata_enabled == 1:
                monthly_bonus_accrual = rupee(basic * bonus_rate)
            else:
                full_month_basic = monthly_salary * 0.40
                monthly_bonus_accrual = rupee(full_month_basic * bonus_rate)

        bonus_ctc = monthly_bonus_accrual

        # Festival bonus payout logic
        festival_bonus = 0

        if festival_bonus_enabled == 1 and month_num == festival_bonus_month:
            cur.execute("""
                SELECT COALESCE(SUM(bonus_ctc), 0) AS accumulated_bonus
                FROM payroll_history
                WHERE company_id = ?
                  AND emp_code = ?
                  AND is_current = 1
                  AND month >= ?
                  AND month <= ?
            """, (
                company_id,
                emp_code,
                financial_year_start,
                month
            ))

            bonus_data = cur.fetchone()
            accumulated_bonus = float(bonus_data["accumulated_bonus"] or 0)

            festival_bonus = rupee(accumulated_bonus + monthly_bonus_accrual)

        festival_bonus = rupee(festival_bonus + attendance_bonus)

        total_deductions = rupee(
            esi_employee
            + professional_tax
            + pf_employee
            + lwf_employee
            + tds
            + manual_deduction
        )

        gratuity = rupee(basic * gratuity_rate)

        total_contributions = rupee(
            esi_employer
            + pf_employer
            + gratuity
            + lwf_employer
        )

        net_pay = rupee(
            gross
            + overtime_amount
            + festival_bonus
            - total_deductions
        )

        monthly_ctc = rupee(
            gross
            + overtime_amount
            + total_contributions
            + bonus_ctc
        )

        annual_ctc = rupee(monthly_ctc * 12)

        cur.execute("""
            INSERT INTO payroll_history (
                company_id,
                emp_code,
                employee_name,
                role,
                department,
                gender,
                month,
                monthly_salary,

                paid_leave_days,
                lwp_days,
                lwp_deduction,
                payable_days,

                basic,
                da,
                hra,
                special_allowance,
                other_allowance,
                gross,

                esi_employee,
                professional_tax,
                pf_employee,
                lwf_employee,
                tds,
                manual_deduction,
                total_deductions,

                esi_employer,
                pf_employer,
                gratuity,
                bonus_ctc,
                festival_bonus,
                lwf_employer,
                total_contributions,

                net_pay,
                monthly_ctc,
                annual_ctc,

                overtime_hours,
                overtime_amount,
                created_at,
                run_id,
                is_current
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            company_id,
            emp_code,
            row["employee_name"],
            row["role"],
            row["department"],
            gender,

            month,
            monthly_salary,

            paid_leave_days,
            lwp_days,
            lwp_deduction,
            payable_days,

            basic,
            da,
            hra,
            special_allowance,
            other_allowance,
            gross,

            esi_employee,
            professional_tax,
            pf_employee,
            lwf_employee,
            tds,
            manual_deduction,
            total_deductions,

            esi_employer,
            pf_employer,
            gratuity,
            bonus_ctc,
            festival_bonus,
            lwf_employer,
            total_contributions,

            net_pay,
            monthly_ctc,
            annual_ctc,

            overtime_hours,
            overtime_amount,
            now,
            run_id,
            1
        ))

    conn.commit()
    conn.close()

    flash("Payroll run completed successfully.")
    return redirect(url_for("payroll_history"))


@app.route("/payroll-history")
@login_required
def payroll_history():
    month = request.args.get("month", "").strip()
    department = request.args.get("department", "").strip()
    search = request.args.get("search", "").strip()

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    query = """
        SELECT 
            *,
            COALESCE(payable_days, 0) AS payable_days,
            COALESCE(paid_leave_days, 0) AS paid_leave_days,
            COALESCE(lwp_days, 0) AS lwp_days,
            COALESCE(lwp_deduction, 0) AS lwp_deduction,

            COALESCE(overtime_amount, 0) AS overtime_amount,
            COALESCE(festival_bonus, 0) AS festival_bonus,
            COALESCE(bonus_ctc, 0) AS bonus_ctc,

            COALESCE(pf_employee, 0) AS pf_employee,
            COALESCE(esi_employee, 0) AS esi_employee,
            COALESCE(professional_tax, 0) AS professional_tax,
            COALESCE(lwf_employee, 0) AS lwf_employee,
            COALESCE(tds, 0) AS tds,
            COALESCE(manual_deduction, 0) AS manual_deduction,

            COALESCE(pf_employer, 0) AS pf_employer,
            COALESCE(esi_employer, 0) AS esi_employer,
            COALESCE(gratuity, 0) AS gratuity,
            COALESCE(lwf_employer, 0) AS lwf_employer,

            COALESCE(total_contributions, 0) AS total_contributions,
            COALESCE(monthly_ctc, 0) AS monthly_ctc,
            COALESCE(annual_ctc, 0) AS annual_ctc
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
    """

    params = [company_id]

    if month:
        query += " AND month = ?"
        params.append(month)

    if department:
        query += " AND department = ?"
        params.append(department)

    if search:
        query += """
            AND (
                emp_code LIKE ?
                OR employee_name LIKE ?
                OR role LIKE ?
                OR department LIKE ?
            )
        """
        search_value = f"%{search}%"
        params.extend([search_value, search_value, search_value, search_value])

    query += " ORDER BY month DESC, id DESC"

    cur.execute(query, tuple(params))
    records = cur.fetchall()

    cur.execute("""
        SELECT DISTINCT department
        FROM employees
        WHERE company_id = ?
          AND department IS NOT NULL
          AND department != ''
        ORDER BY department
    """, (company_id,))
    departments = cur.fetchall()

    total_employees = len(records)

    total_gross = round(sum(float(row["gross"] or 0) for row in records))
    total_overtime_amount = round(sum(float(row["overtime_amount"] or 0) for row in records))
    total_festival_bonus = round(sum(float(row["festival_bonus"] or 0) for row in records))
    total_bonus_ctc = round(sum(float(row["bonus_ctc"] or 0) for row in records))

    total_net_pay = round(sum(float(row["net_pay"] or 0) for row in records))
    total_deductions = round(sum(float(row["total_deductions"] or 0) for row in records))

    total_pf_employee = round(sum(float(row["pf_employee"] or 0) for row in records))
    total_esi_employee = round(sum(float(row["esi_employee"] or 0) for row in records))
    total_professional_tax = round(sum(float(row["professional_tax"] or 0) for row in records))
    total_lwf_employee = round(sum(float(row["lwf_employee"] or 0) for row in records))
    total_tds = round(sum(float(row["tds"] or 0) for row in records))
    total_manual_deduction = round(sum(float(row["manual_deduction"] or 0) for row in records))

    total_pf_employer = round(sum(float(row["pf_employer"] or 0) for row in records))
    total_esi_employer = round(sum(float(row["esi_employer"] or 0) for row in records))
    total_gratuity = round(sum(float(row["gratuity"] or 0) for row in records))
    total_lwf_employer = round(sum(float(row["lwf_employer"] or 0) for row in records))

    total_employer_cost = round(
        total_pf_employer
        + total_esi_employer
        + total_gratuity
        + total_lwf_employer
    )

    total_contributions = round(sum(float(row["total_contributions"] or 0) for row in records))
    total_monthly_ctc = round(sum(float(row["monthly_ctc"] or 0) for row in records))
    total_annual_ctc = round(sum(float(row["annual_ctc"] or 0) for row in records))

    total_paid_leave_days = round(sum(float(row["paid_leave_days"] or 0) for row in records), 2)
    total_lwp_days = round(sum(float(row["lwp_days"] or 0) for row in records), 2)
    total_lwp_deduction = round(sum(float(row["lwp_deduction"] or 0) for row in records))

    conn.close()

    return render_template(
        "payroll_history.html",
        records=records,
        departments=departments,

        selected_month=month,
        selected_department=department,
        search=search,

        total_employees=total_employees,

        total_gross=total_gross,
        total_overtime_amount=total_overtime_amount,
        total_festival_bonus=total_festival_bonus,
        total_bonus_ctc=total_bonus_ctc,

        total_net_pay=total_net_pay,
        total_deductions=total_deductions,

        total_pf_employee=total_pf_employee,
        total_esi_employee=total_esi_employee,
        total_professional_tax=total_professional_tax,
        total_lwf_employee=total_lwf_employee,
        total_tds=total_tds,
        total_manual_deduction=total_manual_deduction,

        total_pf_employer=total_pf_employer,
        total_esi_employer=total_esi_employer,
        total_gratuity=total_gratuity,
        total_lwf_employer=total_lwf_employer,
        total_employer_cost=total_employer_cost,

        total_contributions=total_contributions,
        total_monthly_ctc=total_monthly_ctc,
        total_annual_ctc=total_annual_ctc,

        total_paid_leave_days=total_paid_leave_days,
        total_lwp_days=total_lwp_days,
        total_lwp_deduction=total_lwp_deduction,

        now=datetime.datetime.now()
    )


@app.route("/export-excel")
@login_required
def export_excel():
    if not require_pro_feature("Upgrade to PRO to use Excel export."):
        return redirect(url_for("pricing"))

    month = request.args.get("month", "").strip()
    department = request.args.get("department", "").strip()

    if not month:
        flash("Please select month to export payroll", "warning")
        return redirect(url_for("payroll_history"))

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()

    query = """
        SELECT 
            p.emp_code,
            p.employee_name,
            p.role,
            p.department,
            p.gender,

            COALESCE(e.uan_no, '') AS uan_no,
            COALESCE(e.esic_no, '') AS esic_no,
            COALESCE(e.bank_name, '') AS bank_name,
            COALESCE(e.account_no, '') AS account_no,
            COALESCE(e.ifsc_code, '') AS ifsc_code,

            p.month,
            p.monthly_salary,

            COALESCE(a.working_days, 0) AS attendance_working_days,
            COALESCE(a.present_days, 0) AS attendance_present_days,
            COALESCE(a.weekly_off, 0) AS attendance_weekly_off,
            COALESCE(a.paid_leave, 0) AS attendance_paid_leave,
            COALESCE(a.holiday, 0) AS attendance_holiday,
            COALESCE(a.lop_days, 0) AS attendance_lop_days,
            COALESCE(a.paid_days, 0) AS attendance_paid_days,

            COALESCE(p.payable_days, 0) AS payable_days,
            COALESCE(p.paid_leave_days, 0) AS paid_leave_days,
            COALESCE(p.lwp_days, 0) AS lwp_days,
            COALESCE(p.lwp_deduction, 0) AS lwp_deduction,

            COALESCE(p.basic, 0) AS basic,
            COALESCE(p.da, 0) AS da,
            COALESCE(p.hra, 0) AS hra,
            COALESCE(p.special_allowance, 0) AS special_allowance,
            COALESCE(p.other_allowance, 0) AS other_allowance,
            COALESCE(p.gross, 0) AS gross,

            COALESCE(p.overtime_hours, 0) AS overtime_hours,
            COALESCE(p.overtime_amount, 0) AS overtime_amount,

            COALESCE(p.esi_employee, 0) AS esi_employee,
            COALESCE(p.professional_tax, 0) AS professional_tax,
            COALESCE(p.pf_employee, 0) AS pf_employee,
            COALESCE(p.lwf_employee, 0) AS lwf_employee,
            COALESCE(p.tds, 0) AS tds,
            COALESCE(p.manual_deduction, 0) AS manual_deduction,
            COALESCE(p.total_deductions, 0) AS total_deductions,

            COALESCE(p.esi_employer, 0) AS esi_employer,
            COALESCE(p.pf_employer, 0) AS pf_employer,
            COALESCE(p.gratuity, 0) AS gratuity,
            COALESCE(p.lwf_employer, 0) AS lwf_employer,

            COALESCE(p.bonus_ctc, 0) AS bonus_ctc,
            COALESCE(p.festival_bonus, 0) AS festival_bonus,
            COALESCE(p.total_contributions, 0) AS total_contributions,

            COALESCE(p.net_pay, 0) AS net_pay,
            COALESCE(p.monthly_ctc, 0) AS monthly_ctc,
            COALESCE(p.annual_ctc, 0) AS annual_ctc,

            p.created_at

        FROM payroll_history p

        LEFT JOIN employees e
          ON p.company_id = e.company_id
         AND p.emp_code = e.emp_code

        LEFT JOIN attendance a
          ON p.company_id = a.company_id
         AND p.emp_code = a.emp_code
         AND p.month = a.month

        WHERE p.company_id = ?
          AND p.month = ?
          AND p.is_current = 1
    """

    params = [company_id, month]

    if department:
        query += " AND p.department = ?"
        params.append(department)

    query += " ORDER BY p.emp_code"

    raw_df = pd.read_sql_query(query, conn, params=tuple(params))
    conn.close()

    if raw_df.empty:
        flash("No payroll data found for selected month", "warning")
        return redirect(url_for("payroll_history", month=month))

    raw_df["employer_total"] = (
        raw_df["pf_employer"].fillna(0)
        + raw_df["esi_employer"].fillna(0)
        + raw_df["gratuity"].fillna(0)
        + raw_df["lwf_employer"].fillna(0)
    )

    df = raw_df.copy()

    numeric_columns = [
        "monthly_salary",
        "attendance_working_days",
        "attendance_present_days",
        "attendance_weekly_off",
        "attendance_paid_leave",
        "attendance_holiday",
        "attendance_lop_days",
        "attendance_paid_days",
        "payable_days",
        "paid_leave_days",
        "lwp_days",
        "lwp_deduction",
        "basic",
        "da",
        "hra",
        "special_allowance",
        "other_allowance",
        "gross",
        "overtime_hours",
        "overtime_amount",
        "esi_employee",
        "professional_tax",
        "pf_employee",
        "lwf_employee",
        "tds",
        "manual_deduction",
        "total_deductions",
        "esi_employer",
        "pf_employer",
        "gratuity",
        "lwf_employer",
        "employer_total",
        "bonus_ctc",
        "festival_bonus",
        "total_contributions",
        "net_pay",
        "monthly_ctc",
        "annual_ctc"
    ]

    for col in numeric_columns:
        if col in df.columns:
            df[col] = df[col].fillna(0).round().astype(int)

    text_columns = [
        "emp_code",
        "employee_name",
        "role",
        "department",
        "gender",
        "uan_no",
        "esic_no",
        "bank_name",
        "account_no",
        "ifsc_code",
        "month",
        "created_at"
    ]

    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)

    df = df.rename(columns={
        "emp_code": "Employee Code",
        "employee_name": "Employee Name",
        "role": "Designation",
        "department": "Department",
        "gender": "Gender",

        "uan_no": "UAN No",
        "esic_no": "ESIC No",
        "bank_name": "Bank Name",
        "account_no": "Account No",
        "ifsc_code": "IFSC Code",

        "month": "Month",
        "monthly_salary": "Monthly Salary",

        "attendance_working_days": "Working Days",
        "attendance_present_days": "Present Days",
        "attendance_weekly_off": "Weekly Off",
        "attendance_paid_leave": "Attendance Paid Leave",
        "attendance_holiday": "Holiday",
        "attendance_lop_days": "Attendance LOP Days",
        "attendance_paid_days": "Attendance Paid Days",

        "payable_days": "Final Payable Days",
        "paid_leave_days": "Final Paid Leave Days",
        "lwp_days": "Final LWP Days",
        "lwp_deduction": "LWP Deduction",

        "basic": "Basic",
        "da": "DA",
        "hra": "HRA",
        "special_allowance": "Special Allowance",
        "other_allowance": "Other Allowance",
        "gross": "Gross Salary",

        "overtime_hours": "Overtime Hours",
        "overtime_amount": "Overtime Amount",

        "esi_employee": "ESIC Employee",
        "professional_tax": "Professional Tax",
        "pf_employee": "PF Employee",
        "lwf_employee": "LWF Employee",
        "tds": "TDS",
        "manual_deduction": "Manual Deduction",
        "total_deductions": "Total Deductions",

        "esi_employer": "ESIC Employer",
        "pf_employer": "PF Employer",
        "gratuity": "Gratuity",
        "lwf_employer": "LWF Employer",
        "employer_total": "Employer Total",

        "bonus_ctc": "Bonus CTC",
        "festival_bonus": "Festival Bonus",
        "total_contributions": "Total Contributions",

        "net_pay": "Net Pay",
        "monthly_ctc": "Monthly CTC",
        "annual_ctc": "Annual CTC",

        "created_at": "Created At"
    })

    final_columns = [
        "Employee Code",
        "Employee Name",
        "Designation",
        "Department",
        "Gender",

        "UAN No",
        "ESIC No",
        "Bank Name",
        "Account No",
        "IFSC Code",

        "Month",
        "Monthly Salary",

        "Working Days",
        "Present Days",
        "Weekly Off",
        "Attendance Paid Leave",
        "Holiday",
        "Attendance LOP Days",
        "Attendance Paid Days",

        "Final Payable Days",
        "Final Paid Leave Days",
        "Final LWP Days",
        "LWP Deduction",

        "Basic",
        "DA",
        "HRA",
        "Special Allowance",
        "Other Allowance",
        "Gross Salary",

        "Overtime Hours",
        "Overtime Amount",

        "PF Employee",
        "ESIC Employee",
        "Professional Tax",
        "LWF Employee",
        "TDS",
        "Manual Deduction",
        "Total Deductions",

        "Festival Bonus",
        "Net Pay",

        "PF Employer",
        "ESIC Employer",
        "Gratuity",
        "LWF Employer",
        "Employer Total",

        "Bonus CTC",
        "Total Contributions",
        "Monthly CTC",
        "Annual CTC",

        "Created At"
    ]

    final_columns = [col for col in final_columns if col in df.columns]
    payroll_register_df = df[final_columns]

    # Bank Payment Sheet
    bank_payment_df = raw_df[[
        "emp_code",
        "employee_name",
        "department",
        "bank_name",
        "account_no",
        "ifsc_code",
        "net_pay"
    ]].copy()

    bank_payment_df.insert(0, "Sr No", range(1, len(bank_payment_df) + 1))
    bank_payment_df.insert(1, "Month", month)

    bank_payment_df.rename(columns={
        "emp_code": "Emp Code",
        "employee_name": "Employee Name",
        "department": "Department",
        "bank_name": "Bank Name",
        "account_no": "Account No",
        "ifsc_code": "IFSC Code",
        "net_pay": "Net Pay"
    }, inplace=True)

    bank_payment_df["Net Pay"] = bank_payment_df["Net Pay"].fillna(0).round().astype(int)
    bank_payment_df["Payment Mode"] = "Bank Transfer"
    bank_payment_df["Payment Status"] = "Pending"
    bank_payment_df["Remarks"] = ""

    bank_payment_df = bank_payment_df[[
        "Sr No",
        "Month",
        "Emp Code",
        "Employee Name",
        "Department",
        "Bank Name",
        "Account No",
        "IFSC Code",
        "Net Pay",
        "Payment Mode",
        "Payment Status",
        "Remarks"
    ]]

    total_payment_amount = round(float(bank_payment_df["Net Pay"].fillna(0).sum()))

    total_row = pd.DataFrame([{
        "Sr No": "",
        "Month": "",
        "Emp Code": "",
        "Employee Name": "TOTAL PAYMENT AMOUNT",
        "Department": "",
        "Bank Name": "",
        "Account No": "",
        "IFSC Code": "",
        "Net Pay": total_payment_amount,
        "Payment Mode": "",
        "Payment Status": "",
        "Remarks": ""
    }])

    bank_payment_df = pd.concat([bank_payment_df, total_row], ignore_index=True)

    # Deduction Summary
    deduction_items = [
        ("PF Employee", raw_df["pf_employee"].fillna(0).sum()),
        ("ESIC Employee", raw_df["esi_employee"].fillna(0).sum()),
        ("Professional Tax", raw_df["professional_tax"].fillna(0).sum()),
        ("LWF Employee", raw_df["lwf_employee"].fillna(0).sum()),
        ("TDS", raw_df["tds"].fillna(0).sum()),
        ("Manual Deduction", raw_df["manual_deduction"].fillna(0).sum()),
        ("LWP Deduction", raw_df["lwp_deduction"].fillna(0).sum()),
        ("Total Deductions", raw_df["total_deductions"].fillna(0).sum())
    ]

    deduction_summary_df = pd.DataFrame({
        "Month": [month] * len(deduction_items),
        "Deduction Head": [item[0] for item in deduction_items],
        "Amount": [round(float(item[1])) for item in deduction_items]
    })

    # Employer Contribution Summary
    employer_items = [
        ("PF Employer", raw_df["pf_employer"].fillna(0).sum()),
        ("ESIC Employer", raw_df["esi_employer"].fillna(0).sum()),
        ("Gratuity", raw_df["gratuity"].fillna(0).sum()),
        ("LWF Employer", raw_df["lwf_employer"].fillna(0).sum()),
        ("Employer Total", raw_df["employer_total"].fillna(0).sum()),
        ("Bonus CTC", raw_df["bonus_ctc"].fillna(0).sum()),
        ("Total Contributions", raw_df["total_contributions"].fillna(0).sum()),
        ("Total Monthly CTC", raw_df["monthly_ctc"].fillna(0).sum()),
        ("Total Annual CTC", raw_df["annual_ctc"].fillna(0).sum())
    ]

    employer_summary_df = pd.DataFrame({
        "Month": [month] * len(employer_items),
        "Employer Cost Head": [item[0] for item in employer_items],
        "Amount": [round(float(item[1])) for item in employer_items]
    })

    file_name = f"payroll_{month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        payroll_register_df.to_excel(writer, index=False, sheet_name="Payroll Register")
        bank_payment_df.to_excel(writer, index=False, sheet_name="Bank Payment Sheet")
        deduction_summary_df.to_excel(writer, index=False, sheet_name="Deduction Summary")
        employer_summary_df.to_excel(writer, index=False, sheet_name="Employer Contribution")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        total_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        warning_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(cell_value))

                ws.column_dimensions[column_letter].width = max_length + 3

            # Text formatting
            text_format_headers = [
                "UAN No",
                "ESIC No",
                "Account No",
                "IFSC Code"
            ]

            for col_idx, header_cell in enumerate(ws[1], start=1):
                if header_cell.value in text_format_headers:
                    for row_cells in ws.iter_rows(
                        min_row=2,
                        max_row=ws.max_row,
                        min_col=col_idx,
                        max_col=col_idx
                    ):
                        for cell in row_cells:
                            cell.number_format = "@"

            # Amount formatting by header names
            amount_headers = [
                "Monthly Salary",
                "LWP Deduction",
                "Basic",
                "DA",
                "HRA",
                "Special Allowance",
                "Other Allowance",
                "Gross Salary",
                "Overtime Amount",
                "PF Employee",
                "ESIC Employee",
                "Professional Tax",
                "LWF Employee",
                "TDS",
                "Manual Deduction",
                "Total Deductions",
                "Festival Bonus",
                "Net Pay",
                "PF Employer",
                "ESIC Employer",
                "Gratuity",
                "LWF Employer",
                "Employer Total",
                "Bonus CTC",
                "Total Contributions",
                "Monthly CTC",
                "Annual CTC",
                "Amount"
            ]

            for col_idx, header_cell in enumerate(ws[1], start=1):
                if header_cell.value in amount_headers:
                    for row_cells in ws.iter_rows(
                        min_row=2,
                        max_row=ws.max_row,
                        min_col=col_idx,
                        max_col=col_idx
                    ):
                        for cell in row_cells:
                            cell.number_format = '₹#,##0'

            # Bank Payment Sheet special formatting
            if sheet_name == "Bank Payment Sheet":
                for row_idx in range(2, ws.max_row):
                    bank_name = str(ws.cell(row=row_idx, column=6).value or "").strip()
                    account_no = str(ws.cell(row=row_idx, column=7).value or "").strip()
                    ifsc_code = str(ws.cell(row=row_idx, column=8).value or "").strip()

                    if bank_name == "" or account_no == "" or ifsc_code == "":
                        for col_idx in range(1, ws.max_column + 1):
                            ws.cell(row=row_idx, column=col_idx).fill = warning_fill

                        ws.cell(row=row_idx, column=12).value = "Bank details missing"

                total_row_idx = ws.max_row

                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=total_row_idx, column=col_idx)
                    cell.font = Font(bold=True, color="166534")
                    cell.fill = total_fill

                ws.cell(row=total_row_idx, column=4).value = "TOTAL PAYMENT AMOUNT"
                ws.cell(row=total_row_idx, column=9).number_format = '₹#,##0'

            # Summary sheets total row highlight
            if sheet_name in ["Deduction Summary", "Employer Contribution"]:
                last_row = ws.max_row

                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=last_row, column=col_idx)
                    cell.font = Font(bold=True)

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_name
    )


@app.route("/hr-audit-report")
@login_required
def hr_audit_report():
    if not require_pro_feature("Upgrade to PRO to download HR Audit Report."):
        return redirect(url_for("pricing"))

    month = request.args.get("month", "").strip()

    if not month:
        flash("Please select month for HR Audit Report.")
        return redirect(url_for("payroll_history"))

    company_id = current_company_id()
    conn = get_db()

    employees_df = pd.read_sql_query("""
        SELECT 
            emp_code,
            employee_name,
            role,
            department,
            gender,
            monthly_salary,
            tax_regime,
            other_annual_deductions,
            special_allowance,
            uan_no,
            esic_no,
            bank_name,
            account_no,
            ifsc_code
        FROM employees
        WHERE company_id = ?
        ORDER BY emp_code
    """, conn, params=(company_id,))

    attendance_df = pd.read_sql_query("""
        SELECT 
            a.emp_code,
            e.employee_name,
            e.role,
            e.department,
            a.month,
            a.working_days,
            a.present_days,
            COALESCE(a.weekly_off, 0) AS weekly_off,
            COALESCE(a.paid_leave, 0) AS paid_leave,
            COALESCE(a.holiday, 0) AS holiday,
            COALESCE(a.lop_days, 0) AS lop_days,
            COALESCE(a.paid_days, 0) AS paid_days,
            COALESCE(a.overtime_hours, 0) AS overtime_hours,
            COALESCE(a.bonus, 0) AS attendance_bonus,
            COALESCE(a.manual_deduction, 0) AS manual_deduction
        FROM attendance a
        LEFT JOIN employees e
          ON a.company_id = e.company_id
         AND a.emp_code = e.emp_code
        WHERE a.company_id = ?
          AND a.month = ?
        ORDER BY a.emp_code
    """, conn, params=(company_id, month))

    payroll_df = pd.read_sql_query("""
        SELECT 
            emp_code,
            employee_name,
            role,
            department,
            gender,
            month,

            COALESCE(payable_days, 0) AS payable_days,
            COALESCE(paid_leave_days, 0) AS paid_leave_days,
            COALESCE(lwp_days, 0) AS lwp_days,
            COALESCE(lwp_deduction, 0) AS lwp_deduction,

            COALESCE(basic, 0) AS basic,
            COALESCE(da, 0) AS da,
            COALESCE(hra, 0) AS hra,
            COALESCE(special_allowance, 0) AS special_allowance,
            COALESCE(other_allowance, 0) AS other_allowance,
            COALESCE(gross, 0) AS gross,

            COALESCE(overtime_hours, 0) AS overtime_hours,
            COALESCE(overtime_amount, 0) AS overtime_amount,

            COALESCE(pf_employee, 0) AS pf_employee,
            COALESCE(esi_employee, 0) AS esi_employee,
            COALESCE(professional_tax, 0) AS professional_tax,
            COALESCE(lwf_employee, 0) AS lwf_employee,
            COALESCE(tds, 0) AS tds,
            COALESCE(manual_deduction, 0) AS manual_deduction,
            COALESCE(total_deductions, 0) AS total_deductions,

            COALESCE(pf_employer, 0) AS pf_employer,
            COALESCE(esi_employer, 0) AS esi_employer,
            COALESCE(gratuity, 0) AS gratuity,
            COALESCE(lwf_employer, 0) AS lwf_employer,

            COALESCE(bonus_ctc, 0) AS bonus_ctc,
            COALESCE(festival_bonus, 0) AS festival_bonus,
            COALESCE(total_contributions, 0) AS total_contributions,

            COALESCE(net_pay, 0) AS net_pay,
            COALESCE(monthly_ctc, 0) AS monthly_ctc,
            COALESCE(annual_ctc, 0) AS annual_ctc,

            created_at
        FROM payroll_history
        WHERE company_id = ?
          AND month = ?
          AND is_current = 1
        ORDER BY emp_code
    """, conn, params=(company_id, month))

    leave_df = pd.read_sql_query("""
        SELECT
            lr.emp_code,
            e.employee_name,
            e.department,
            lr.leave_type,
            lr.start_date,
            lr.end_date,
            lr.total_days,
            lr.status,
            lr.reason
        FROM leave_requests lr
        LEFT JOIN employees e
          ON lr.company_id = e.company_id
         AND lr.emp_code = e.emp_code
        WHERE lr.company_id = ?
          AND (
                substr(lr.start_date, 1, 7) = ?
                OR substr(lr.end_date, 1, 7) = ?
          )
        ORDER BY lr.emp_code, lr.start_date
    """, conn, params=(company_id, month, month))

    conn.close()

    if employees_df.empty:
        flash("No employee data found for HR audit.")
        return redirect(url_for("employees_list"))

    def is_missing(value):
        if pd.isna(value):
            return True
        value = str(value).strip()
        return value == "" or value.lower() in ["nan", "none", "null", "-"]

    employees_df["uan_status"] = employees_df["uan_no"].apply(lambda x: "Missing" if is_missing(x) else "OK")
    employees_df["esic_status"] = employees_df["esic_no"].apply(lambda x: "Missing" if is_missing(x) else "OK")
    employees_df["bank_status"] = employees_df["bank_name"].apply(lambda x: "Missing" if is_missing(x) else "OK")
    employees_df["account_status"] = employees_df["account_no"].apply(lambda x: "Missing" if is_missing(x) else "OK")
    employees_df["ifsc_status"] = employees_df["ifsc_code"].apply(lambda x: "Missing" if is_missing(x) else "OK")
    employees_df["department_status"] = employees_df["department"].apply(lambda x: "Missing" if is_missing(x) else "OK")
    employees_df["gender_status"] = employees_df["gender"].apply(lambda x: "Missing" if is_missing(x) else "OK")

    employees_df["salary_status"] = employees_df["monthly_salary"].apply(
        lambda x: "Invalid" if pd.isna(x) or float(x or 0) <= 0 else "OK"
    )

    attendance_emp_codes = set(attendance_df["emp_code"].astype(str)) if not attendance_df.empty else set()
    payroll_emp_codes = set(payroll_df["emp_code"].astype(str)) if not payroll_df.empty else set()

    employees_df["attendance_status"] = employees_df["emp_code"].astype(str).apply(
        lambda x: "Attendance Missing" if x not in attendance_emp_codes else "OK"
    )

    employees_df["payroll_status"] = employees_df["emp_code"].astype(str).apply(
        lambda x: "Payroll Missing" if x not in payroll_emp_codes else "OK"
    )

    employees_df["audit_month"] = month

    audit_columns = [
        "audit_month",
        "emp_code",
        "employee_name",
        "role",
        "department",
        "gender",
        "monthly_salary",
        "uan_no",
        "esic_no",
        "bank_name",
        "account_no",
        "ifsc_code",
        "uan_status",
        "esic_status",
        "bank_status",
        "account_status",
        "ifsc_status",
        "department_status",
        "gender_status",
        "salary_status",
        "attendance_status",
        "payroll_status"
    ]

    employee_audit_df = employees_df[audit_columns]

    if attendance_df.empty:
        total_working_days = 0
        total_present_days = 0
        total_weekly_off = 0
        total_attendance_paid_leave = 0
        total_holiday = 0
        total_attendance_lop_days = 0
        total_attendance_paid_days = 0
        total_overtime_hours = 0
        total_attendance_bonus = 0
        total_attendance_manual_deduction = 0
    else:
        total_working_days = round(float(attendance_df["working_days"].fillna(0).sum()), 2)
        total_present_days = round(float(attendance_df["present_days"].fillna(0).sum()), 2)
        total_weekly_off = round(float(attendance_df["weekly_off"].fillna(0).sum()), 2)
        total_attendance_paid_leave = round(float(attendance_df["paid_leave"].fillna(0).sum()), 2)
        total_holiday = round(float(attendance_df["holiday"].fillna(0).sum()), 2)
        total_attendance_lop_days = round(float(attendance_df["lop_days"].fillna(0).sum()), 2)
        total_attendance_paid_days = round(float(attendance_df["paid_days"].fillna(0).sum()), 2)
        total_overtime_hours = round(float(attendance_df["overtime_hours"].fillna(0).sum()), 2)
        total_attendance_bonus = round(float(attendance_df["attendance_bonus"].fillna(0).sum()))
        total_attendance_manual_deduction = round(float(attendance_df["manual_deduction"].fillna(0).sum()))

    if payroll_df.empty:
        total_payable_days = 0
        total_paid_leave_days = 0
        total_lwp_days = 0
        total_lwp_deduction = 0
        employees_with_lwp = 0

        total_gross = 0
        total_overtime_amount = 0
        total_festival_bonus = 0
        total_bonus_ctc = 0
        total_pf_employee = 0
        total_esi_employee = 0
        total_professional_tax = 0
        total_lwf_employee = 0
        total_tds = 0
        total_manual_deduction = 0
        total_deductions = 0

        total_pf_employer = 0
        total_esi_employer = 0
        total_gratuity = 0
        total_lwf_employer = 0
        total_employer_contribution = 0
        total_contributions = 0

        total_net_pay = 0
        total_monthly_ctc = 0
        total_annual_ctc = 0
    else:
        payroll_df["employer_total"] = (
            payroll_df["pf_employer"].fillna(0)
            + payroll_df["esi_employer"].fillna(0)
            + payroll_df["gratuity"].fillna(0)
            + payroll_df["lwf_employer"].fillna(0)
        )

        total_payable_days = round(float(payroll_df["payable_days"].fillna(0).sum()), 2)
        total_paid_leave_days = round(float(payroll_df["paid_leave_days"].fillna(0).sum()), 2)
        total_lwp_days = round(float(payroll_df["lwp_days"].fillna(0).sum()), 2)
        total_lwp_deduction = round(float(payroll_df["lwp_deduction"].fillna(0).sum()))
        employees_with_lwp = int((payroll_df["lwp_days"].fillna(0) > 0).sum())

        total_gross = round(float(payroll_df["gross"].fillna(0).sum()))
        total_overtime_amount = round(float(payroll_df["overtime_amount"].fillna(0).sum()))
        total_festival_bonus = round(float(payroll_df["festival_bonus"].fillna(0).sum()))
        total_bonus_ctc = round(float(payroll_df["bonus_ctc"].fillna(0).sum()))

        total_pf_employee = round(float(payroll_df["pf_employee"].fillna(0).sum()))
        total_esi_employee = round(float(payroll_df["esi_employee"].fillna(0).sum()))
        total_professional_tax = round(float(payroll_df["professional_tax"].fillna(0).sum()))
        total_lwf_employee = round(float(payroll_df["lwf_employee"].fillna(0).sum()))
        total_tds = round(float(payroll_df["tds"].fillna(0).sum()))
        total_manual_deduction = round(float(payroll_df["manual_deduction"].fillna(0).sum()))
        total_deductions = round(float(payroll_df["total_deductions"].fillna(0).sum()))

        total_pf_employer = round(float(payroll_df["pf_employer"].fillna(0).sum()))
        total_esi_employer = round(float(payroll_df["esi_employer"].fillna(0).sum()))
        total_gratuity = round(float(payroll_df["gratuity"].fillna(0).sum()))
        total_lwf_employer = round(float(payroll_df["lwf_employer"].fillna(0).sum()))
        total_employer_contribution = round(float(payroll_df["employer_total"].fillna(0).sum()))
        total_contributions = round(float(payroll_df["total_contributions"].fillna(0).sum()))

        total_net_pay = round(float(payroll_df["net_pay"].fillna(0).sum()))
        total_monthly_ctc = round(float(payroll_df["monthly_ctc"].fillna(0).sum()))
        total_annual_ctc = round(float(payroll_df["annual_ctc"].fillna(0).sum()))

    approved_leave_count = 0
    rejected_leave_count = 0
    pending_leave_count = 0

    if not leave_df.empty:
        approved_leave_count = int((leave_df["status"] == "Approved").sum())
        rejected_leave_count = int((leave_df["status"] == "Rejected").sum())
        pending_leave_count = int((leave_df["status"] == "Pending").sum())

    summary_items = [
        ("Total Employees", len(employees_df)),
        ("Attendance Uploaded", len(attendance_emp_codes)),
        ("Payroll Processed", len(payroll_emp_codes)),

        ("Total Working Days", total_working_days),
        ("Total Present Days", total_present_days),
        ("Total Weekly Off", total_weekly_off),
        ("Total Attendance Paid Leave", total_attendance_paid_leave),
        ("Total Holiday", total_holiday),
        ("Total Attendance LOP Days", total_attendance_lop_days),
        ("Total Attendance Paid Days", total_attendance_paid_days),
        ("Total Overtime Hours", total_overtime_hours),
        ("Total Attendance Bonus", total_attendance_bonus),
        ("Total Attendance Manual Deduction", total_attendance_manual_deduction),

        ("Total Final Payable Days", total_payable_days),
        ("Total Final Paid Leave Days", total_paid_leave_days),
        ("Total Final LWP Days", total_lwp_days),
        ("Total LWP Deduction", total_lwp_deduction),
        ("Employees With LWP", employees_with_lwp),

        ("Total Gross Salary", total_gross),
        ("Total Overtime Amount", total_overtime_amount),
        ("Total Festival Bonus Paid", total_festival_bonus),
        ("Total Bonus CTC", total_bonus_ctc),

        ("Total PF Employee", total_pf_employee),
        ("Total ESIC Employee", total_esi_employee),
        ("Total Professional Tax", total_professional_tax),
        ("Total LWF Employee", total_lwf_employee),
        ("Total TDS", total_tds),
        ("Total Manual Deduction", total_manual_deduction),
        ("Total Deductions", total_deductions),

        ("Total PF Employer", total_pf_employer),
        ("Total ESIC Employer", total_esi_employer),
        ("Total Gratuity", total_gratuity),
        ("Total LWF Employer", total_lwf_employer),
        ("Total Employer Contribution", total_employer_contribution),
        ("Total Contributions", total_contributions),

        ("Total Net Pay", total_net_pay),
        ("Total Monthly CTC", total_monthly_ctc),
        ("Total Annual CTC", total_annual_ctc),

        ("Approved Leave Requests", approved_leave_count),
        ("Rejected Leave Requests", rejected_leave_count),
        ("Pending Leave Requests", pending_leave_count),

        ("Missing UAN", int((employees_df["uan_status"] == "Missing").sum())),
        ("Missing ESIC No.", int((employees_df["esic_status"] == "Missing").sum())),
        ("Missing Bank Name", int((employees_df["bank_status"] == "Missing").sum())),
        ("Missing Account No.", int((employees_df["account_status"] == "Missing").sum())),
        ("Missing IFSC", int((employees_df["ifsc_status"] == "Missing").sum())),
        ("Missing Department", int((employees_df["department_status"] == "Missing").sum())),
        ("Missing Gender", int((employees_df["gender_status"] == "Missing").sum())),
        ("Invalid Salary", int((employees_df["salary_status"] == "Invalid").sum())),
        ("Attendance Missing", int((employees_df["attendance_status"] == "Attendance Missing").sum())),
        ("Payroll Missing", int((employees_df["payroll_status"] == "Payroll Missing").sum()))
    ]

    summary_df = pd.DataFrame({
        "Audit Month": [month] * len(summary_items),
        "Audit Item": [item[0] for item in summary_items],
        "Count / Amount": [item[1] for item in summary_items]
    })

    payroll_missing_df = employee_audit_df[
        employee_audit_df["payroll_status"] == "Payroll Missing"
    ]

    attendance_missing_df = employee_audit_df[
        employee_audit_df["attendance_status"] == "Attendance Missing"
    ]

    # Bank Payment Sheet
    # Payroll processed employees ke net pay ko bank payment format me export karega.
    if payroll_df.empty:
        bank_payment_df = pd.DataFrame(columns=[
            "Sr No",
            "Audit Month",
            "Emp Code",
            "Employee Name",
            "Department",
            "Bank Name",
            "Account No",
            "IFSC Code",
            "Net Pay",
            "Payment Mode",
            "Payment Status",
            "Remarks"
        ])
        bank_payment_total = 0

    else:
        bank_master_df = employees_df[[
            "emp_code",
            "bank_name",
            "account_no",
            "ifsc_code"
        ]].copy()

        bank_payment_df = payroll_df[[
            "emp_code",
            "employee_name",
            "department",
            "net_pay"
        ]].copy()

        bank_payment_df = bank_payment_df.merge(
            bank_master_df,
            on="emp_code",
            how="left"
        )

        bank_payment_df.insert(0, "Sr No", range(1, len(bank_payment_df) + 1))
        bank_payment_df.insert(1, "Audit Month", month)

        bank_payment_df.rename(columns={
            "emp_code": "Emp Code",
            "employee_name": "Employee Name",
            "department": "Department",
            "bank_name": "Bank Name",
            "account_no": "Account No",
            "ifsc_code": "IFSC Code",
            "net_pay": "Net Pay"
        }, inplace=True)

        bank_payment_df["Payment Mode"] = "Bank Transfer"
        bank_payment_df["Payment Status"] = "Pending"
        bank_payment_df["Remarks"] = ""

        bank_payment_df = bank_payment_df[[
            "Sr No",
            "Audit Month",
            "Emp Code",
            "Employee Name",
            "Department",
            "Bank Name",
            "Account No",
            "IFSC Code",
            "Net Pay",
            "Payment Mode",
            "Payment Status",
            "Remarks"
        ]]

        bank_payment_total = round(float(bank_payment_df["Net Pay"].fillna(0).sum()))

        # Total payment row at bottom
        total_row = pd.DataFrame([{
            "Sr No": "",
            "Audit Month": "",
            "Emp Code": "",
            "Employee Name": "TOTAL PAYMENT AMOUNT",
            "Department": "",
            "Bank Name": "",
            "Account No": "",
            "IFSC Code": "",
            "Net Pay": bank_payment_total,
            "Payment Mode": "",
            "Payment Status": "",
            "Remarks": ""
        }])

        bank_payment_df = pd.concat(
            [bank_payment_df, total_row],
            ignore_index=True
        )

    # Add bank payment total in audit summary also
    summary_df = pd.concat([
        summary_df,
        pd.DataFrame({
            "Audit Month": [month],
            "Audit Item": ["Bank Payment Total Amount"],
            "Count / Amount": [bank_payment_total]
        })
    ], ignore_index=True)

    file_name = f"hr_audit_report_{month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Audit Summary")
        employee_audit_df.to_excel(writer, index=False, sheet_name="Employee Master Audit")
        attendance_df.to_excel(writer, index=False, sheet_name="Attendance Audit")
        payroll_df.to_excel(writer, index=False, sheet_name="Payroll Audit")
        leave_df.to_excel(writer, index=False, sheet_name="Leave Audit")
        bank_payment_df.to_excel(writer, index=False, sheet_name="Bank Payment Sheet")
        attendance_missing_df.to_excel(writer, index=False, sheet_name="Attendance Missing")
        payroll_missing_df.to_excel(writer, index=False, sheet_name="Payroll Missing")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        missing_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        ok_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        total_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(cell_value))

                    if cell.row != 1:
                        if cell_value in ["Missing", "Invalid", "Attendance Missing", "Payroll Missing", "Pending"]:
                            cell.fill = missing_fill
                        elif cell_value == "OK":
                            cell.fill = ok_fill

                ws.column_dimensions[column_letter].width = max_length + 3

            # Bank Payment Sheet special formatting
            if sheet_name == "Bank Payment Sheet":
                # Net Pay amount formatting
                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row=row_idx, column=9).number_format = '₹#,##0'

                # Keep bank account number and IFSC as text
                for col in ["G", "H"]:
                    for cell in ws[col]:
                        cell.number_format = "@"

                # Highlight total row
                total_row_idx = ws.max_row

                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=total_row_idx, column=col_idx)
                    cell.font = Font(bold=True, color="166534")
                    cell.fill = total_fill

                ws.cell(row=total_row_idx, column=4).value = "TOTAL PAYMENT AMOUNT"
                ws.cell(row=total_row_idx, column=9).number_format = '₹#,##0'

            # Audit Summary amount formatting
            if sheet_name == "Audit Summary":
                for row_idx in range(2, ws.max_row + 1):
                    item_name = str(ws.cell(row=row_idx, column=2).value or "")

                    if "Amount" in item_name or "Salary" in item_name or "Pay" in item_name or "CTC" in item_name or "Deduction" in item_name:
                        ws.cell(row=row_idx, column=3).number_format = '₹#,##0'

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_name
    )


@app.route("/download-bank-payment-sheet")
@login_required
def download_bank_payment_sheet():
    if not require_pro_feature("Upgrade to PRO to download Bank Payment Sheet."):
        return redirect(url_for("pricing"))

    month = request.args.get("month", "").strip()

    if not month:
        flash("Please select month to download Bank Payment Sheet.", "warning")
        return redirect(url_for("payroll_history"))

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()

    payroll_df = pd.read_sql_query("""
        SELECT 
            p.emp_code,
            p.employee_name,
            p.department,
            COALESCE(p.net_pay, 0) AS net_pay,

            COALESCE(e.bank_name, '') AS bank_name,
            COALESCE(e.account_no, '') AS account_no,
            COALESCE(e.ifsc_code, '') AS ifsc_code

        FROM payroll_history p
        LEFT JOIN employees e
          ON p.company_id = e.company_id
         AND p.emp_code = e.emp_code

        WHERE p.company_id = ?
          AND p.month = ?
          AND p.is_current = 1

        ORDER BY p.emp_code
    """, conn, params=(company_id, month))

    conn.close()

    if payroll_df.empty:
        flash("No payroll data found for selected month.", "warning")
        return redirect(url_for("payroll_history", month=month))

    bank_payment_df = payroll_df.copy()

    bank_payment_df.insert(0, "Sr No", range(1, len(bank_payment_df) + 1))
    bank_payment_df.insert(1, "Month", month)

    bank_payment_df.rename(columns={
        "emp_code": "Emp Code",
        "employee_name": "Employee Name",
        "department": "Department",
        "bank_name": "Bank Name",
        "account_no": "Account No",
        "ifsc_code": "IFSC Code",
        "net_pay": "Net Pay"
    }, inplace=True)

    bank_payment_df["Payment Mode"] = "Bank Transfer"
    bank_payment_df["Payment Status"] = "Pending"
    bank_payment_df["Remarks"] = ""

    bank_payment_df = bank_payment_df[[
        "Sr No",
        "Month",
        "Emp Code",
        "Employee Name",
        "Department",
        "Bank Name",
        "Account No",
        "IFSC Code",
        "Net Pay",
        "Payment Mode",
        "Payment Status",
        "Remarks"
    ]]

    total_payment = round(float(bank_payment_df["Net Pay"].fillna(0).sum()))

    total_row = pd.DataFrame([{
        "Sr No": "",
        "Month": "",
        "Emp Code": "",
        "Employee Name": "TOTAL PAYMENT AMOUNT",
        "Department": "",
        "Bank Name": "",
        "Account No": "",
        "IFSC Code": "",
        "Net Pay": total_payment,
        "Payment Mode": "",
        "Payment Status": "",
        "Remarks": ""
    }])

    bank_payment_df = pd.concat([bank_payment_df, total_row], ignore_index=True)

    file_name = f"bank_payment_sheet_{month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        bank_payment_df.to_excel(writer, index=False, sheet_name="Bank Payment Sheet")

        workbook = writer.book
        ws = workbook["Bank Payment Sheet"]

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        total_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        warning_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0")
        )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row_cells:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

                if str(cell.value or "").strip() in ["Pending", "Missing", ""]:
                    pass

        # Net Pay amount formatting
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=9).number_format = '₹#,##0'

        # Keep Account No and IFSC as text
        for col in ["G", "H"]:
            for cell in ws[col]:
                cell.number_format = "@"

        # Highlight missing bank details
        for row_idx in range(2, ws.max_row):
            bank_name = str(ws.cell(row=row_idx, column=6).value or "").strip()
            account_no = str(ws.cell(row=row_idx, column=7).value or "").strip()
            ifsc_code = str(ws.cell(row=row_idx, column=8).value or "").strip()

            if bank_name == "" or account_no == "" or ifsc_code == "":
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = warning_fill

                ws.cell(row=row_idx, column=12).value = "Bank details missing"

        # Highlight total row
        total_row_idx = ws.max_row

        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=total_row_idx, column=col_idx)
            cell.font = Font(bold=True, color="166534")
            cell.fill = total_fill
            cell.border = thin_border

        ws.cell(row=total_row_idx, column=4).value = "TOTAL PAYMENT AMOUNT"
        ws.cell(row=total_row_idx, column=9).number_format = '₹#,##0'

        # Auto width
        for column_cells in ws.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                cell_value = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(cell_value))

            ws.column_dimensions[column_letter].width = max_length + 4

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_name
    )

@app.route("/full-and-final", methods=["GET", "POST"])
@login_required
def full_and_final():
    if not require_pro_feature("Upgrade to PRO to use Full & Final Settlement."):
        return redirect(url_for("pricing"))

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    def to_float(value, default=0):
        try:
            if value is None or str(value).strip() == "":
                return default
            return float(value)
        except Exception:
            return default

    def money_round(value):
        try:
            return round(float(value or 0))
        except Exception:
            return 0

    def clean_text(value, default="-"):
        value = str(value or "").strip()
        return value if value else default

    def clean_reason(value):
        value = str(value or "").strip()
        if not value:
            return "-"
        return value[:1].upper() + value[1:].lower()

    conn = get_db()
    cur = conn.cursor()

    if request.method == "POST":
        try:
            emp_code = request.form.get("emp_code", "").strip()
            last_working_day = request.form.get("last_working_day", "").strip()
            settlement_month = request.form.get("settlement_month", "").strip()

            reason = clean_reason(request.form.get("reason", ""))
            remarks = clean_text(request.form.get("remarks", ""), "-")

            bonus_payable = to_float(request.form.get("bonus_payable"), 0)
            gratuity_payable = to_float(request.form.get("gratuity_payable"), 0)
            other_earnings = to_float(request.form.get("other_earnings"), 0)

            notice_recovery = to_float(request.form.get("notice_recovery"), 0)
            loan_recovery = to_float(request.form.get("loan_recovery"), 0)
            advance_recovery = to_float(request.form.get("advance_recovery"), 0)
            other_deductions = to_float(request.form.get("other_deductions"), 0)

            apply_leave_encashment = request.form.get(
                "apply_leave_encashment", "no"
            ).strip().lower()

            errors = []

            if not emp_code:
                errors.append("Employee is required.")

            if not last_working_day:
                errors.append("Last Working Day is required.")

            if not settlement_month:
                errors.append("Settlement Month is required.")

            last_working_date_obj = None

            if last_working_day:
                try:
                    last_working_date_obj = datetime.datetime.strptime(
                        last_working_day, "%Y-%m-%d"
                    )
                except Exception:
                    errors.append("Last Working Day must be a valid date.")

            if settlement_month:
                try:
                    datetime.datetime.strptime(settlement_month, "%Y-%m")
                except Exception:
                    errors.append("Settlement Month must be in YYYY-MM format.")

            if last_working_date_obj and settlement_month:
                last_working_month = last_working_date_obj.strftime("%Y-%m")
                if last_working_month != settlement_month:
                    errors.append("Settlement Month must match the Last Working Day month.")

            money_fields = {
                "Bonus Payable": bonus_payable,
                "Gratuity Payable": gratuity_payable,
                "Other Earnings": other_earnings,
                "Notice Recovery": notice_recovery,
                "Loan Recovery": loan_recovery,
                "Advance Recovery": advance_recovery,
                "Other Deductions": other_deductions
            }

            for label, value in money_fields.items():
                if value < 0:
                    errors.append(f"{label} cannot be negative.")

                if value > 100000000:
                    errors.append(f"{label} amount is too high. Please check.")

            if apply_leave_encashment not in ["yes", "no"]:
                apply_leave_encashment = "no"

            if errors:
                flash(" ".join(errors), "danger")
                return redirect(url_for("full_and_final"))

            cur.execute("""
                SELECT 
                    emp_code,
                    employee_name,
                    role,
                    department,
                    monthly_salary
                FROM employees
                WHERE company_id = ?
                  AND emp_code = ?
                LIMIT 1
            """, (company_id, emp_code))

            emp = cur.fetchone()

            if not emp:
                flash("Employee not found.", "danger")
                return redirect(url_for("full_and_final"))

            monthly_salary = to_float(emp["monthly_salary"], 0)

            if monthly_salary <= 0:
                flash(
                    "Employee monthly salary is invalid. Please update employee master.",
                    "danger"
                )
                return redirect(url_for("full_and_final"))

            cur.execute("""
                SELECT id
                FROM full_final_settlements
                WHERE company_id = ?
                  AND emp_code = ?
                  AND settlement_month = ?
                LIMIT 1
            """, (company_id, emp_code, settlement_month))

            existing_settlement = cur.fetchone()

            if existing_settlement:
                flash(
                    "Full & Final settlement already exists for this employee and settlement month. "
                    "Please delete the existing settlement before creating a new one.",
                    "warning"
                )
                return redirect(url_for("full_and_final"))

            # F&F salary basis:
            # Earned salary is calculated on fixed 30-day basis.
            # Example: salary / 30 * paid days.
            paid_days = last_working_date_obj.day if last_working_date_obj else 0
            paid_days = max(0, min(paid_days, 30))

            per_day_salary = monthly_salary / 30
            earned_salary = money_round(per_day_salary * paid_days)

            # Leave encashment:
            # Currently only paid_leave balance is considered.
            cur.execute("""
                SELECT COALESCE(paid_leave, 0) AS paid_leave
                FROM leave_balances
                WHERE company_id = ?
                  AND emp_code = ?
                LIMIT 1
            """, (company_id, emp_code))

            leave_row = cur.fetchone()

            if leave_row:
                leave_balance = to_float(leave_row["paid_leave"], 0)
            else:
                leave_balance = 0

            leave_balance = max(0, leave_balance)

            if apply_leave_encashment == "yes":
                leave_encashment = money_round(per_day_salary * leave_balance)
            else:
                leave_encashment = 0

            total_earnings = money_round(
                earned_salary
                + bonus_payable
                + gratuity_payable
                + other_earnings
                + leave_encashment
            )

            total_deductions = money_round(
                notice_recovery
                + loan_recovery
                + advance_recovery
                + other_deductions
            )

            final_payable = money_round(total_earnings - total_deductions)

            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cur.execute("""
                INSERT INTO full_final_settlements (
                    company_id,
                    emp_code,
                    employee_name,
                    role,
                    department,
                    last_working_day,
                    settlement_month,
                    monthly_salary,
                    paid_days,
                    earned_salary,
                    leave_balance,
                    leave_encashment,
                    bonus_payable,
                    gratuity_payable,
                    other_earnings,
                    notice_recovery,
                    loan_recovery,
                    advance_recovery,
                    other_deductions,
                    total_earnings,
                    total_deductions,
                    final_payable,
                    reason,
                    remarks,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                company_id,
                emp["emp_code"],
                emp["employee_name"],
                emp["role"],
                emp["department"],

                last_working_day,
                settlement_month,

                monthly_salary,
                paid_days,
                earned_salary,

                leave_balance,
                leave_encashment,
                bonus_payable,
                gratuity_payable,
                other_earnings,

                notice_recovery,
                loan_recovery,
                advance_recovery,
                other_deductions,

                total_earnings,
                total_deductions,
                final_payable,

                reason,
                remarks,
                now
            ))

            conn.commit()

            flash(
                f"Full & Final settlement created successfully for {emp['employee_name']}. "
                f"Final Payable: ₹{final_payable}",
                "success"
            )

            return redirect(url_for("full_and_final"))

        except Exception as e:
            conn.rollback()
            flash(f"Error while creating Full & Final settlement: {str(e)}", "danger")
            return redirect(url_for("full_and_final"))

        finally:
            try:
                conn.close()
            except Exception:
                pass

    try:
        cur.execute("""
            SELECT 
                e.emp_code,
                e.employee_name,
                e.role,
                e.department,
                e.monthly_salary,
                COALESCE(lb.paid_leave, 0) AS paid_leave
            FROM employees e
            LEFT JOIN leave_balances lb
              ON e.company_id = lb.company_id
             AND e.emp_code = lb.emp_code
            WHERE e.company_id = ?
            ORDER BY e.employee_name
        """, (company_id,))

        employees = cur.fetchall()

        cur.execute("""
            SELECT *
            FROM full_final_settlements
            WHERE company_id = ?
            ORDER BY id DESC
        """, (company_id,))

        settlements = cur.fetchall()

        total_settlements = len(settlements)
        total_final_payable = money_round(
            sum(float(row["final_payable"] or 0) for row in settlements)
        )
        total_earnings = money_round(
            sum(float(row["total_earnings"] or 0) for row in settlements)
        )
        total_deductions = money_round(
            sum(float(row["total_deductions"] or 0) for row in settlements)
        )

    except Exception as e:
        conn.rollback()
        flash(f"Error while loading Full & Final data: {str(e)}", "danger")
        return redirect(url_for("dashboard"))

    finally:
        try:
            conn.close()
        except Exception:
            pass

    return render_template(
        "full_and_final.html",
        employees=employees,
        settlements=settlements,
        total_settlements=total_settlements,
        total_final_payable=total_final_payable,
        total_earnings=total_earnings,
        total_deductions=total_deductions
    )


@app.route("/download-fnf-excel/<int:settlement_id>")
@login_required
def download_fnf_excel(settlement_id):
    if not require_pro_feature("Upgrade to PRO to download Full & Final Excel."):
        return redirect(url_for("pricing"))

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            f.*,
            COALESCE(c.company_name, '') AS company_name,
            COALESCE(c.address, '') AS company_address,
            COALESCE(c.email, '') AS company_email,
            COALESCE(c.phone, '') AS company_phone
        FROM full_final_settlements f
        JOIN companies c
          ON f.company_id = c.id
        WHERE f.id = ?
          AND f.company_id = ?
    """, (settlement_id, company_id))

    row = cur.fetchone()
    conn.close()

    if not row:
        flash("Full & Final settlement not found.", "warning")
        return redirect(url_for("full_and_final"))

    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    def money(value):
        try:
            return round(float(value or 0))
        except Exception:
            return 0

    def number_value(value):
        try:
            num = float(value or 0)
            if num.is_integer():
                return int(num)
            return round(num, 2)
        except Exception:
            return 0

    def clean(value, default="-"):
        if value is None:
            return default

        value = str(value).strip()

        if value == "" or value.lower() in ["nan", "none", "null"]:
            return default

        return value

    def clean_reason(value):
        value = clean(value, "-")

        if value == "-":
            return "-"

        return value[:1].upper() + value[1:].lower()

    def clean_company_name(value):
        value = clean(value, "SMARTHIRE AI")

        # Fix rare spacing issue like "SMARTHIRE A I"
        value = value.replace("A I", "AI").replace("SmartHire A I", "SmartHire AI")
        value = value.replace("SMART HIRE A I", "SMART HIRE AI")

        return value

    def safe_filename(value, default="file"):
        value = clean(value, default)

        for ch in [" ", "/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
            value = value.replace(ch, "_")

        return value

    emp_code = safe_filename(row["emp_code"], "employee")
    settlement_month = safe_filename(row["settlement_month"], "month")

    file_name = f"fnf_settlement_{emp_code}_{settlement_month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    company_phone = clean(row["company_phone"], "")

    # Keep leading zero in Excel phone number.
    if company_phone and company_phone != "-":
        company_phone = str(company_phone)

    data = [
        ["FULL & FINAL SETTLEMENT STATEMENT", ""],
        ["", ""],

        ["Company Details", ""],
        ["Company Name", clean_company_name(row["company_name"])],
        ["Company Address", clean(row["company_address"])],
        ["Company Email", clean(row["company_email"])],
        ["Company Phone", company_phone if company_phone else "-"],
        ["", ""],

        ["Employee Details", ""],
        ["Employee Code", clean(row["emp_code"])],
        ["Employee Name", clean(row["employee_name"])],
        ["Designation", clean(row["role"])],
        ["Department", clean(row["department"])],
        ["Last Working Day", clean(row["last_working_day"])],
        ["Settlement Month", clean(row["settlement_month"])],
        ["Reason for Leaving", clean_reason(row["reason"])],
        ["", ""],

        ["Earnings", "Amount"],
        ["Monthly Salary", money(row["monthly_salary"])],
        ["Paid Days", number_value(row["paid_days"])],
        ["Earned Salary", money(row["earned_salary"])],
        ["Leave Balance", number_value(row["leave_balance"])],
        ["Leave Encashment", money(row["leave_encashment"])],
        ["Bonus Payable", money(row["bonus_payable"])],
        ["Gratuity Payable", money(row["gratuity_payable"])],
        ["Other Earnings", money(row["other_earnings"])],
        ["Total Earnings", money(row["total_earnings"])],
        ["", ""],

        ["Deductions", "Amount"],
        ["Notice Recovery", money(row["notice_recovery"])],
        ["Loan Recovery", money(row["loan_recovery"])],
        ["Advance Recovery", money(row["advance_recovery"])],
        ["Other Deductions", money(row["other_deductions"])],
        ["Total Deductions", money(row["total_deductions"])],
        ["", ""],

        ["Final Settlement", ""],
        ["Final Payable", money(row["final_payable"])],
        ["Remarks", clean(row["remarks"])],
        ["Created At", clean(row["created_at"])],
        ["", ""],

        ["Approvals", ""],
        ["Prepared By", "________________________"],
        ["Checked By", "________________________"],
        ["HR / Authorized Signatory", "________________________"],
    ]

    df = pd.DataFrame(data, columns=["Particulars", "Details"])

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="F&F Settlement")

        wb = writer.book
        ws = wb["F&F Settlement"]

        # Black & white professional styling
        black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        light_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        white_font = Font(bold=True, color="FFFFFF")
        title_font = Font(bold=True, size=14, color="000000")
        section_font = Font(bold=True, color="000000")
        bold_font = Font(bold=True, color="000000")
        normal_font = Font(color="000000")

        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000")
        )

        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        right = Alignment(horizontal="right", vertical="center", wrap_text=True)

        ws.freeze_panes = "A4"

        # Pandas header row
        for cell in ws[1]:
            cell.fill = black_fill
            cell.font = white_font
            cell.alignment = center
            cell.border = thin_border

        # Main title row
        ws.merge_cells("A2:B2")
        ws["A2"].font = title_font
        ws["A2"].alignment = center
        ws["A2"].border = thin_border
        ws["A2"].fill = white_fill
        ws.row_dimensions[2].height = 26

        section_labels = [
            "Company Details",
            "Employee Details",
            "Earnings",
            "Deductions",
            "Final Settlement",
            "Approvals"
        ]

        amount_labels = [
            "Monthly Salary",
            "Earned Salary",
            "Leave Encashment",
            "Bonus Payable",
            "Gratuity Payable",
            "Other Earnings",
            "Total Earnings",
            "Notice Recovery",
            "Loan Recovery",
            "Advance Recovery",
            "Other Deductions",
            "Total Deductions",
            "Final Payable"
        ]

        text_labels = [
            "Company Phone",
            "Employee Code",
            "Last Working Day",
            "Settlement Month",
            "Created At"
        ]

        bold_rows = [
            "Total Earnings",
            "Total Deductions",
            "Final Payable"
        ]

        for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
            label = row_cells[0].value

            for cell in row_cells:
                cell.border = thin_border
                cell.font = normal_font
                cell.fill = white_fill
                cell.alignment = left

            if label in section_labels:
                ws.merge_cells(
                    start_row=row_cells[0].row,
                    start_column=1,
                    end_row=row_cells[0].row,
                    end_column=2
                )

                section_cell = ws.cell(row=row_cells[0].row, column=1)
                section_cell.font = section_font
                section_cell.fill = light_fill
                section_cell.alignment = center
                section_cell.border = thin_border

            if label in ["Earnings", "Deductions"]:
                ws.cell(row=row_cells[0].row, column=1).font = section_font
                ws.cell(row=row_cells[0].row, column=2).font = section_font
                ws.cell(row=row_cells[0].row, column=1).alignment = center
                ws.cell(row=row_cells[0].row, column=2).alignment = center

            if label in bold_rows:
                for cell in row_cells:
                    cell.font = bold_font
                    cell.fill = light_fill

            if label == "Final Payable":
                for cell in row_cells:
                    cell.font = Font(bold=True, size=12, color="000000")
                    cell.fill = white_fill

            if label in amount_labels:
                amount_cell = ws.cell(row=row_cells[0].row, column=2)
                amount_cell.number_format = '₹#,##0'
                amount_cell.alignment = right

            if label in text_labels:
                text_cell = ws.cell(row=row_cells[0].row, column=2)
                text_cell.number_format = "@"
                text_cell.alignment = left

        # Ensure phone remains text and leading zero is visible
        for row_idx in range(1, ws.max_row + 1):
            if ws.cell(row=row_idx, column=1).value == "Company Phone":
                ws.cell(row=row_idx, column=2).value = str(company_phone) if company_phone else "-"
                ws.cell(row=row_idx, column=2).number_format = "@"

        ws.column_dimensions["A"].width = 32
        ws.column_dimensions["B"].width = 45

        for row_idx in range(1, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 20

        # Page setup for clean printing
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        ws.sheet_properties.pageSetUpPr.fitToPage = True

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_name
    )


@app.route("/download-fnf-pdf/<int:settlement_id>")
@login_required
def download_fnf_pdf(settlement_id):
    if not require_pro_feature("Upgrade to PRO to download Full & Final PDF."):
        return redirect(url_for("pricing"))

    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            f.*,
            COALESCE(c.company_name, '') AS company_name,
            COALESCE(c.address, '') AS company_address,
            COALESCE(c.email, '') AS company_email,
            COALESCE(c.phone, '') AS company_phone
        FROM full_final_settlements f
        JOIN companies c
          ON f.company_id = c.id
        WHERE f.id = ?
          AND f.company_id = ?
    """, (settlement_id, company_id))

    row = cur.fetchone()
    conn.close()

    if not row:
        flash("Full & Final settlement not found.", "warning")
        return redirect(url_for("full_and_final"))

    os.makedirs(PAYSLIP_FOLDER, exist_ok=True)

    def clean(value, default="-"):
        if value is None:
            return default

        value = str(value).strip()

        if value == "" or value.lower() in ["nan", "none", "null"]:
            return default

        return value

    def clean_reason(value):
        value = clean(value, "-")

        if value == "-":
            return "-"

        return value[:1].upper() + value[1:].lower()

    def clean_company_name(value):
        value = clean(value, "SMARTHIRE AI")

        # Fix rare spacing issue like "SMARTHIRE A I"
        value = value.replace("A I", "AI")
        value = value.replace("SMART HIRE A I", "SMART HIRE AI")
        value = value.replace("SmartHire A I", "SmartHire AI")

        return value

    def safe_filename(value, default="file"):
        value = clean(value, default)

        for ch in [" ", "/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
            value = value.replace(ch, "_")

        return value

    def short_text(value, max_len=70):
        value = clean(value, "")

        if len(value) > max_len:
            return value[:max_len - 3] + "..."

        return value

    def money(value):
        try:
            return f"Rs. {int(round(float(value or 0)))}"
        except Exception:
            return "Rs. 0"

    def number_text(value, default="0"):
        try:
            if value is None or str(value).strip() == "":
                return default

            num = float(value)

            if num.is_integer():
                return str(int(num))

            return str(round(num, 2))
        except Exception:
            return default

    def row_value(key, default=""):
        try:
            return row[key]
        except Exception:
            return default

    emp_code = safe_filename(row_value("emp_code"), "employee")
    settlement_month = safe_filename(row_value("settlement_month"), "month")

    file_name = f"fnf_settlement_{emp_code}_{settlement_month}.pdf"
    file_path = os.path.join(PAYSLIP_FOLDER, file_name)

    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter

    company_name = clean_company_name(row_value("company_name"))
    company_address = clean(row_value("company_address"), "")
    company_email = clean(row_value("company_email"), "")
    company_phone = clean(row_value("company_phone"), "")

    y = height - 32

    # Header box - black & white
    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(1, 1, 1)
    c.rect(35, y - 72, width - 70, 72, fill=1, stroke=1)

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 16)

    # Avoid forced weird spacing. Keep company name clean.
    c.drawCentredString(width / 2, y - 18, short_text(company_name.upper(), 60))

    c.setFont("Helvetica-Bold", 8.5)
    c.drawCentredString(width / 2, y - 34, "Full & Final Settlement Statement")

    c.setFont("Helvetica", 8.2)

    if company_address not in ["", "-"]:
        c.drawCentredString(width / 2, y - 50, short_text(company_address, 90))

    email_phone_parts = []

    if company_email not in ["", "-"]:
        email_phone_parts.append(company_email)

    if company_phone not in ["", "-"]:
        email_phone_parts.append(str(company_phone))

    email_phone_line = " | ".join(email_phone_parts)

    if email_phone_line:
        c.drawCentredString(width / 2, y - 64, short_text(email_phone_line, 90))

    y -= 96

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(
        width / 2,
        y,
        f"FULL & FINAL SETTLEMENT - {clean(row_value('settlement_month'))}"
    )

    y -= 26

    # Employee details
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(40, y, "Employee & Exit Details")

    y -= 9

    details_box_height = 116
    c.rect(40, y - details_box_height, width - 80, details_box_height)

    # Center divider
    c.line(width / 2, y, width / 2, y - details_box_height)

    details_left = [
        ("Employee Code", row_value("emp_code")),
        ("Employee Name", row_value("employee_name")),
        ("Designation", row_value("role")),
        ("Department", row_value("department")),
        ("Monthly Salary", money(row_value("monthly_salary", 0))),
    ]

    details_right = [
        ("Last Working Day", row_value("last_working_day")),
        ("Settlement Month", row_value("settlement_month")),
        ("Reason", clean_reason(row_value("reason"))),
        ("Paid Days", number_text(row_value("paid_days", 0))),
        ("Created At", row_value("created_at")),
    ]

    yy = y - 22

    for label, value in details_left:
        c.setFont("Helvetica-Bold", 8.2)
        c.drawString(55, yy, f"{label}:")
        c.setFont("Helvetica", 8.2)
        c.drawString(155, yy, short_text(clean(value), 30))
        yy -= 19

    yy = y - 22

    for label, value in details_right:
        c.setFont("Helvetica-Bold", 8.2)
        c.drawString(320, yy, f"{label}:")
        c.setFont("Helvetica", 8.2)
        c.drawString(430, yy, short_text(clean(value), 28))
        yy -= 19

    y -= 138

    # Settlement calculation
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(40, y, "Settlement Calculation")

    y -= 13

    x = 40
    table_w = width - 80
    row_h = 19
    table_y = y
    rows_count = 10

    c.rect(x, table_y - row_h * rows_count, table_w, row_h * rows_count)

    # Vertical lines
    c.line(x + table_w * 0.35, table_y, x + table_w * 0.35, table_y - row_h * rows_count)
    c.line(x + table_w * 0.50, table_y, x + table_w * 0.50, table_y - row_h * rows_count)
    c.line(x + table_w * 0.78, table_y, x + table_w * 0.78, table_y - row_h * rows_count)

    # Horizontal lines
    for i in range(rows_count + 1):
        c.line(x, table_y - row_h * i, x + table_w, table_y - row_h * i)

    # Header row
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + 8, table_y - 13, "Earnings")
    c.drawRightString(x + table_w * 0.50 - 8, table_y - 13, "Amount")
    c.drawString(x + table_w * 0.50 + 8, table_y - 13, "Deductions")
    c.drawRightString(x + table_w - 8, table_y - 13, "Amount")

    earnings = [
        ("Monthly Salary", row_value("monthly_salary", 0), "money"),
        ("Paid Days", row_value("paid_days", 0), "number"),
        ("Earned Salary", row_value("earned_salary", 0), "money"),
        ("Leave Balance", row_value("leave_balance", 0), "number"),
        ("Leave Encashment", row_value("leave_encashment", 0), "money"),
        ("Bonus Payable", row_value("bonus_payable", 0), "money"),
        ("Gratuity Payable", row_value("gratuity_payable", 0), "money"),
        ("Other Earnings", row_value("other_earnings", 0), "money"),
        ("Total Earnings", row_value("total_earnings", 0), "money"),
    ]

    deductions = [
        ("Notice Recovery", row_value("notice_recovery", 0)),
        ("Loan Recovery", row_value("loan_recovery", 0)),
        ("Advance Recovery", row_value("advance_recovery", 0)),
        ("Other Deductions", row_value("other_deductions", 0)),
        ("Total Deductions", row_value("total_deductions", 0)),
        ("", ""),
        ("", ""),
        ("", ""),
        ("", ""),
    ]

    start_y = table_y - row_h - 13

    for i in range(9):
        yy = start_y - row_h * i

        earning_label, earning_value, earning_type = earnings[i]

        if earning_label == "Total Earnings":
            c.setFont("Helvetica-Bold", 8.2)
        else:
            c.setFont("Helvetica", 8.2)

        c.drawString(x + 8, yy, earning_label)

        if earning_type == "number":
            c.drawRightString(x + table_w * 0.50 - 8, yy, number_text(earning_value))
        else:
            c.drawRightString(x + table_w * 0.50 - 8, yy, money(earning_value))

        deduction_label, deduction_value = deductions[i]

        if deduction_label:
            if deduction_label == "Total Deductions":
                c.setFont("Helvetica-Bold", 8.2)
            else:
                c.setFont("Helvetica", 8.2)

            c.drawString(x + table_w * 0.50 + 8, yy, deduction_label)
            c.drawRightString(x + table_w - 8, yy, money(deduction_value))

    y = table_y - row_h * rows_count - 24

    # Final payable box
    c.rect(40, y - 32, width - 80, 32)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(55, y - 21, f"FINAL PAYABLE: {money(row_value('final_payable', 0))}")

    y -= 54

    # Remarks
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(40, y, "Remarks")

    y -= 9
    c.rect(40, y - 34, width - 80, 34)

    c.setFont("Helvetica", 8.2)
    remarks = clean(row_value("remarks"), "-")

    if len(remarks) > 115:
        remarks = remarks[:112] + "..."

    c.drawString(55, y - 21, remarks)

    y -= 58

    # Signatures
    c.setFont("Helvetica", 8.2)

    c.line(40, y + 16, 140, y + 16)
    c.drawString(40, y, "Prepared By")

    c.line(260, y + 16, 360, y + 16)
    c.drawString(260, y, "Checked By")

    c.line(470, y + 16, 570, y + 16)
    c.drawString(470, y, "HR / Authorized Signatory")

    y -= 24

    c.setFont("Helvetica-Oblique", 7)
    c.drawCentredString(
        width / 2,
        y,
        "Computer-generated Full & Final settlement. Signature not required if digitally approved."
    )

    c.save()

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_name
    )


@app.route("/delete-fnf/<int:settlement_id>", methods=["POST"])
@login_required
def delete_fnf(settlement_id):
    if not require_pro_feature("Upgrade to PRO to delete Full & Final settlements."):
        return redirect(url_for("pricing"))

    company_id = current_company_id()

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT id, employee_name
            FROM full_final_settlements
            WHERE id = ?
              AND company_id = ?
        """, (settlement_id, company_id))

        row = cur.fetchone()

        if not row:
            flash("Full & Final settlement not found.", "warning")
            return redirect(url_for("full_and_final"))

        cur.execute("""
            DELETE FROM full_final_settlements
            WHERE id = ?
              AND company_id = ?
        """, (settlement_id, company_id))

        deleted_count = cur.rowcount
        conn.commit()

        if deleted_count > 0:
            flash(
                f"Full & Final settlement deleted successfully for {row['employee_name']}.",
                "success"
            )
        else:
            flash("Full & Final settlement not found or already deleted.", "warning")

    except Exception as e:
        conn.rollback()
        flash(f"Error while deleting Full & Final settlement: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("full_and_final"))


@app.route("/delete-payroll/<month>", methods=["POST"])
@login_required
def delete_payroll(month):
    company_id = current_company_id()

    if not company_id:
        flash("Company not found. Please login again.", "danger")
        return redirect(url_for("login"))

    # Month validation: expected format YYYY-MM
    try:
        datetime.datetime.strptime(month, "%Y-%m")
    except ValueError:
        flash("Invalid payroll month selected.", "danger")
        return redirect(url_for("payroll_history"))

    conn = get_db()
    cur = conn.cursor()

    try:
        # Check records before delete
        cur.execute("""
            SELECT COUNT(*) AS total_records
            FROM payroll_history
            WHERE company_id = ?
              AND month = ?
        """, (company_id, month))

        result = cur.fetchone()
        total_records = result["total_records"] if result else 0

        if total_records <= 0:
            flash(f"No payroll records found for {month}.", "warning")
            return redirect(url_for("payroll_history", month=month))

        # Delete payroll records for selected company + selected month only.
        # Leave requests and leave balances are intentionally untouched.
        cur.execute("""
            DELETE FROM payroll_history
            WHERE company_id = ?
              AND month = ?
        """, (company_id, month))

        deleted_count = cur.rowcount
        conn.commit()

        flash(
            f"Payroll deleted successfully for {month}. "
            f"{deleted_count} record(s) removed. "
            "Leave records and leave balances were not changed.",
            "success"
        )

    except Exception as e:
        conn.rollback()
        flash(f"Error while deleting payroll: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("payroll_history", month=month))


# ---------------------------
# PAYSLIP PDF
# ---------------------------
def generate_payslip(row):
    os.makedirs(PAYSLIP_FOLDER, exist_ok=True)

    def clean_value(value, default="-"):
        if value is None:
            return default

        value = str(value).strip()

        if value == "":
            return default

        if value.lower() in ["nan", "none", "null"]:
            return default

        return value

    def safe_filename(value, default="file"):
        value = clean_value(value, default)

        for ch in [" ", "/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
            value = value.replace(ch, "_")

        return value

    def to_float(value, default=0):
        try:
            if value is None or str(value).strip() == "":
                return default
            return float(value)
        except Exception:
            return default

    def money(value):
        try:
            return f"Rs. {int(round(float(value or 0)))}"
        except Exception:
            return "Rs. 0"

    def number_text(value, default="0"):
        try:
            if value is None or str(value).strip() == "":
                return default

            num = float(value)

            if num.is_integer():
                return str(int(num))

            return str(round(num, 2))

        except Exception:
            return default

    def row_value(key, default=""):
        try:
            return row[key]
        except Exception:
            return default

    def short_text(value, max_len=55):
        value = clean_value(value, "")

        if len(value) > max_len:
            return value[:max_len - 3] + "..."

        return value

    emp_code_for_file = safe_filename(row_value("emp_code"), "employee")
    month_for_file = safe_filename(row_value("month"), "month")

    file_name = f"{emp_code_for_file}_{month_for_file}.pdf"
    file_path = os.path.join(PAYSLIP_FOLDER, file_name)

    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter

    # Force white page background for official print-friendly PDF
    c.setFillColorRGB(1, 1, 1)
    c.rect(0, 0, width, height, fill=1, stroke=0)

    # Default text and border color black
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

    company_name = clean_value(row_value("company_name"), "SMART HIRE AI PAYROLL")
    company_address = clean_value(row_value("company_address"), "")
    company_email = clean_value(row_value("company_email"), "")
    company_phone = clean_value(row_value("company_phone"), "")

    emp_code = clean_value(row_value("emp_code"))
    employee_name = clean_value(row_value("employee_name"))
    role = clean_value(row_value("role"))
    department = clean_value(row_value("department"))
    gender = clean_value(row_value("gender"))

    uan_no = clean_value(row_value("uan_no"))
    esic_no = clean_value(row_value("esic_no"))
    bank_name = clean_value(row_value("bank_name"))
    account_no = clean_value(row_value("account_no"))
    ifsc_code = clean_value(row_value("ifsc_code"))

    working_days = row_value("attendance_working_days", row_value("working_days", 0))
    present_days = row_value("attendance_present_days", row_value("present_days", 0))
    weekly_off = row_value("attendance_weekly_off", row_value("weekly_off", 0))
    attendance_paid_leave = row_value("attendance_paid_leave", row_value("paid_leave_days", 0))
    holiday = row_value("attendance_holiday", row_value("holiday", 0))
    lop_days = row_value("attendance_lop_days", row_value("lwp_days", 0))
    attendance_paid_days = row_value("attendance_paid_days", row_value("payable_days", 0))
    overtime_hours = row_value("attendance_overtime_hours", row_value("overtime_hours", 0))

    payable_days = row_value("payable_days", attendance_paid_days)
    lwp_deduction = row_value("lwp_deduction", 0)

    calculated_paid_days = (
        to_float(present_days)
        + to_float(weekly_off)
        + to_float(attendance_paid_leave)
        + to_float(holiday)
        - to_float(lop_days)
    )

    if to_float(payable_days) <= 0:
        payable_days = calculated_paid_days

    absent_days = (
        to_float(working_days)
        - to_float(present_days)
        - to_float(weekly_off)
        - to_float(attendance_paid_leave)
        - to_float(holiday)
    )

    if absent_days < 0:
        absent_days = 0

    y = height - 30

    # Header
    c.setFillColorRGB(0.10, 0.17, 0.28)
    c.rect(35, y - 66, width - 70, 66, fill=1, stroke=0)

    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, y - 18, short_text(company_name.upper(), 60))

    c.setFont("Helvetica", 8.5)
    c.drawCentredString(width / 2, y - 34, "Salary Slip")

    address_line = short_text(company_address, 82) if company_address not in ["", "-"] else ""

    email_phone_parts = []

    if company_email not in ["", "-"]:
        email_phone_parts.append(company_email)

    if company_phone not in ["", "-"]:
        email_phone_parts.append(company_phone)

    email_phone_line = " | ".join(email_phone_parts)

    if address_line:
        c.drawCentredString(width / 2, y - 49, address_line)

    if email_phone_line:
        c.drawCentredString(width / 2, y - 61, short_text(email_phone_line, 82))

    y -= 88

    # Title
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width / 2, y, f"SALARY SLIP - {clean_value(row_value('month'))}")

    y -= 26

    # Employee Info
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(40, y, "Employee Information")

    y -= 9

    emp_box_height = 118
    c.rect(40, y - emp_box_height, width - 80, emp_box_height)

    # Center divider
    c.line(width / 2, y, width / 2, y - emp_box_height)

    left = [
        ("Employee Code", emp_code),
        ("Employee Name", employee_name),
        ("Designation", role),
        ("Department", department),
        ("Gender", gender),
    ]

    right = [
        ("UAN No.", uan_no),
        ("ESIC No.", esic_no),
        ("Bank Name", bank_name),
        ("Account No.", account_no),
        ("IFSC Code", ifsc_code),
    ]

    yy = y - 22

    for label, value in left:
        c.setFont("Helvetica-Bold", 8.3)
        c.drawString(55, yy, f"{label}:")
        c.setFont("Helvetica", 8.3)
        c.drawString(150, yy, short_text(clean_value(value), 30))
        yy -= 20

    yy = y - 22

    for label, value in right:
        c.setFont("Helvetica-Bold", 8.3)
        c.drawString(320, yy, f"{label}:")
        c.setFont("Helvetica", 8.3)
        c.drawString(405, yy, short_text(clean_value(value), 28))
        yy -= 20

    y -= 140

    # Attendance Summary
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(40, y, "Attendance & Leave Summary")

    y -= 9

    att_box_height = 66
    c.rect(40, y - att_box_height, width - 80, att_box_height)

    c.setFont("Helvetica-Bold", 7.2)
    c.drawString(50, y - 14, "Working")
    c.drawString(110, y - 14, "Present")
    c.drawString(170, y - 14, "Weekly Off")
    c.drawString(245, y - 14, "Paid Leave")
    c.drawString(320, y - 14, "Holiday")
    c.drawString(380, y - 14, "LOP")
    c.drawString(430, y - 14, "Absent")
    c.drawString(490, y - 14, "Paid Days")

    c.setFont("Helvetica", 7.8)
    c.drawString(50, y - 29, number_text(working_days))
    c.drawString(110, y - 29, number_text(present_days))
    c.drawString(170, y - 29, number_text(weekly_off))
    c.drawString(245, y - 29, number_text(attendance_paid_leave))
    c.drawString(320, y - 29, number_text(holiday))
    c.drawString(380, y - 29, number_text(lop_days))
    c.drawString(430, y - 29, number_text(absent_days))
    c.drawString(490, y - 29, number_text(payable_days))

    c.setFont("Helvetica-Bold", 7.8)
    c.drawString(55, y - 51, "Overtime Hours")
    c.drawString(230, y - 51, "LOP Deduction")
    c.drawString(420, y - 51, "Pay Month")

    c.setFont("Helvetica", 7.8)
    c.drawString(155, y - 51, number_text(overtime_hours))
    c.drawString(330, y - 51, money(lwp_deduction))
    c.drawString(495, y - 51, clean_value(row_value("month")))

    y -= 84

    # Salary Details
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(40, y, "Salary Details")

    y -= 13

    x = 40
    table_w = width - 80
    row_h = 18
    table_y = y
    rows_count = 9

    c.rect(x, table_y - (row_h * rows_count), table_w, row_h * rows_count)

    c.setFillColorRGB(0.12, 0.36, 0.85)
    c.rect(x, table_y - row_h, table_w, row_h, fill=1, stroke=0)

    c.setFillColorRGB(0, 0, 0)
    c.line(x + table_w * 0.28, table_y, x + table_w * 0.28, table_y - row_h * rows_count)
    c.line(x + table_w * 0.50, table_y, x + table_w * 0.50, table_y - row_h * rows_count)
    c.line(x + table_w * 0.78, table_y, x + table_w * 0.78, table_y - row_h * rows_count)

    for i in range(rows_count + 1):
        c.line(x, table_y - row_h * i, x + table_w, table_y - row_h * i)

    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + 8, table_y - 12.5, "Earnings")
    c.drawString(x + table_w * 0.28 + 8, table_y - 12.5, "Amount")
    c.drawString(x + table_w * 0.50 + 8, table_y - 12.5, "Deductions")
    c.drawString(x + table_w * 0.78 + 8, table_y - 12.5, "Amount")

    c.setFillColorRGB(0, 0, 0)

    earnings = [
        ("Basic", row_value("basic", 0)),
        ("DA", row_value("da", 0)),
        ("HRA", row_value("hra", 0)),
        ("Special Allowance", row_value("special_allowance", 0)),
        ("Other Allowance", row_value("other_allowance", 0)),
        ("Overtime Amount", row_value("overtime_amount", 0)),
        ("Festival Bonus", row_value("festival_bonus", 0)),
    ]

    deductions = [
        ("PF Employee", row_value("pf_employee", 0)),
        ("ESIC Employee", row_value("esi_employee", 0)),
        ("Professional Tax", row_value("professional_tax", 0)),
        ("TDS", row_value("tds", 0)),
        ("Manual Deduction", row_value("manual_deduction", 0)),
        ("LWF Employee", row_value("lwf_employee", 0)),
        ("Total Deductions", row_value("total_deductions", 0)),
    ]

    c.setFont("Helvetica", 8.1)
    start_y = table_y - row_h - 12.5

    for i in range(7):
        yy = start_y - row_h * i

        c.drawString(x + 8, yy, earnings[i][0])
        c.drawRightString(x + table_w * 0.50 - 8, yy, money(earnings[i][1]))

        c.drawString(x + table_w * 0.50 + 8, yy, deductions[i][0])
        c.drawRightString(x + table_w - 8, yy, money(deductions[i][1]))

    total_y = table_y - row_h * 8 - 12.5

    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x + 8, total_y, "Gross Earnings")
    c.drawRightString(x + table_w * 0.50 - 8, total_y, money(row_value("gross", 0)))

    c.drawString(x + table_w * 0.50 + 8, total_y, "Net Pay")
    c.drawRightString(x + table_w - 8, total_y, money(row_value("net_pay", 0)))

    y = table_y - row_h * rows_count - 22

    # Employer Contributions
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(40, y, "Employer Contributions / CTC")

    y -= 9

    emp_ctc_box_height = 64
    c.rect(40, y - emp_ctc_box_height, width - 80, emp_ctc_box_height)

    pf_employer = to_float(row_value("pf_employer", 0))
    esi_employer = to_float(row_value("esi_employer", 0))
    gratuity = to_float(row_value("gratuity", 0))
    lwf_employer = to_float(row_value("lwf_employer", 0))

    employer_total = pf_employer + esi_employer + gratuity + lwf_employer

    c.setFont("Helvetica", 8.1)
    c.drawString(55, y - 14, f"PF Employer: {money(pf_employer)}")
    c.drawString(55, y - 29, f"ESIC Employer: {money(esi_employer)}")
    c.drawString(55, y - 44, f"Gratuity: {money(gratuity)}")
    c.drawString(55, y - 59, f"LWF Employer: {money(lwf_employer)}")

    c.drawString(330, y - 14, f"Monthly CTC: {money(row_value('monthly_ctc', 0))}")
    c.drawString(330, y - 29, f"Annual CTC: {money(row_value('annual_ctc', 0))}")
    c.drawString(330, y - 44, f"Bonus CTC: {money(row_value('bonus_ctc', 0))}")

    c.setFont("Helvetica-Bold", 8.1)
    c.drawString(330, y - 59, f"Employer Total: {money(employer_total)}")

    y -= 78

    # Net Pay Highlight
    c.setFillColorRGB(0.86, 0.96, 0.89)
    c.rect(40, y - 32, width - 80, 32, fill=1, stroke=0)

    c.setFillColorRGB(0, 0.35, 0.15)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(55, y - 21, f"NET PAYABLE: {money(row_value('net_pay', 0))}")

    c.setFillColorRGB(0, 0, 0)

    # Signature
    y -= 44
    c.setFont("Helvetica", 8)
    c.drawString(40, y, "Prepared By")
    c.drawString(260, y, "Checked By")
    c.drawString(470, y, "HR / Authorized Signatory")

    c.line(40, y + 16, 140, y + 16)
    c.line(260, y + 16, 360, y + 16)
    c.line(470, y + 16, 570, y + 16)

    y -= 22
    c.setFont("Helvetica-Oblique", 7)
    c.drawCentredString(
        width / 2,
        y,
        "Computer-generated payslip. Signature not required if digitally approved."
    )

    c.save()

    return file_path


@app.route("/download-payslip/<int:payroll_id>")
@login_required
def download_payslip(payroll_id):
    if not require_pro_feature("Upgrade to PRO to download PDF payslips."):
        return redirect(url_for("pricing"))

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            p.*,

            COALESCE(p.paid_leave_days, 0) AS paid_leave_days,
            COALESCE(p.lwp_days, 0) AS lwp_days,
            COALESCE(p.lwp_deduction, 0) AS lwp_deduction,
            COALESCE(p.payable_days, 0) AS payable_days,

            COALESCE(c.company_name, '') AS company_name,
            COALESCE(c.address, '') AS company_address,
            COALESCE(c.email, '') AS company_email,
            COALESCE(c.phone, '') AS company_phone,

            COALESCE(e.uan_no, '') AS uan_no,
            COALESCE(e.esic_no, '') AS esic_no,
            COALESCE(e.bank_name, '') AS bank_name,
            COALESCE(e.account_no, '') AS account_no,
            COALESCE(e.ifsc_code, '') AS ifsc_code,

            COALESCE(a.working_days, 0) AS attendance_working_days,
            COALESCE(a.present_days, 0) AS attendance_present_days,
            COALESCE(a.weekly_off, 0) AS attendance_weekly_off,
            COALESCE(a.paid_leave, 0) AS attendance_paid_leave,
            COALESCE(a.holiday, 0) AS attendance_holiday,
            COALESCE(a.lop_days, 0) AS attendance_lop_days,
            COALESCE(a.paid_days, 0) AS attendance_paid_days,
            COALESCE(a.overtime_hours, 0) AS attendance_overtime_hours

        FROM payroll_history p

        JOIN companies c
          ON p.company_id = c.id

        JOIN employees e
          ON p.company_id = e.company_id
         AND p.emp_code = e.emp_code

        LEFT JOIN attendance a
          ON p.company_id = a.company_id
         AND p.emp_code = a.emp_code
         AND p.month = a.month

        WHERE p.id = ?
          AND p.company_id = ?

        ORDER BY a.id DESC
        LIMIT 1
    """, (payroll_id, current_company_id()))

    row = cur.fetchone()
    conn.close()

    if not row:
        flash("Payslip not found.")
        return redirect(url_for("payroll_history"))

    pdf_path = generate_payslip(row)
    return send_file(pdf_path, as_attachment=True)


@app.route("/download-all-payslips")
@login_required
def download_all_payslips():
    if not require_pro_feature("Upgrade to PRO to download all payslips."):
        return redirect(url_for("pricing"))

    month = request.args.get("month", "").strip()

    if not month:
        flash("Please select month to download payslips.", "warning")
        return redirect(url_for("payroll_history"))

    # Basic month validation: expected YYYY-MM
    try:
        datetime.datetime.strptime(month, "%Y-%m")
    except ValueError:
        flash("Invalid payroll month selected.", "danger")
        return redirect(url_for("payroll_history"))

    company_id = current_company_id()

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            p.*,

            COALESCE(p.paid_leave_days, 0) AS paid_leave_days,
            COALESCE(p.lwp_days, 0) AS lwp_days,
            COALESCE(p.lwp_deduction, 0) AS lwp_deduction,
            COALESCE(p.payable_days, 0) AS payable_days,

            COALESCE(c.company_name, '') AS company_name,
            COALESCE(c.address, '') AS company_address,
            COALESCE(c.email, '') AS company_email,
            COALESCE(c.phone, '') AS company_phone,

            COALESCE(e.uan_no, '') AS uan_no,
            COALESCE(e.esic_no, '') AS esic_no,
            COALESCE(e.bank_name, '') AS bank_name,
            COALESCE(e.account_no, '') AS account_no,
            COALESCE(e.ifsc_code, '') AS ifsc_code,

            COALESCE(a.working_days, 0) AS attendance_working_days,
            COALESCE(a.present_days, 0) AS attendance_present_days,
            COALESCE(a.weekly_off, 0) AS attendance_weekly_off,
            COALESCE(a.paid_leave, 0) AS attendance_paid_leave,
            COALESCE(a.holiday, 0) AS attendance_holiday,
            COALESCE(a.lop_days, 0) AS attendance_lop_days,
            COALESCE(a.paid_days, 0) AS attendance_paid_days,
            COALESCE(a.overtime_hours, 0) AS attendance_overtime_hours

        FROM payroll_history p

        JOIN companies c
          ON p.company_id = c.id

        JOIN employees e
          ON p.company_id = e.company_id
         AND p.emp_code = e.emp_code

        LEFT JOIN attendance a
          ON p.company_id = a.company_id
         AND p.emp_code = a.emp_code
         AND p.month = a.month

        WHERE p.company_id = ?
          AND p.month = ?
          AND p.is_current = 1

        ORDER BY p.emp_code
    """, (company_id, month))

    rows = cur.fetchall()
    conn.close()

    if not rows:
        flash("No payroll data found for selected month.", "warning")
        return redirect(url_for("payroll_history", month=month))

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for row in rows:
            pdf_path = generate_payslip(row)

            emp_code = str(row["emp_code"] or "EMP").strip()
            employee_name = str(row["employee_name"] or "Employee").strip()

            safe_employee_name = (
                employee_name
                .replace(" ", "_")
                .replace("/", "_")
                .replace("\\", "_")
                .replace(":", "_")
                .replace("*", "_")
                .replace("?", "_")
                .replace('"', "_")
                .replace("<", "_")
                .replace(">", "_")
                .replace("|", "_")
            )

            pdf_name = f"{emp_code}_{safe_employee_name}_{month}.pdf"
            zip_file.write(pdf_path, pdf_name)

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=f"payslips_{month}.zip",
        mimetype="application/zip"
    )


@app.route("/faq")
@login_required
def faq():
    return render_template("faq.html")


# ---------------------------
# START APP
# ---------------------------
def setup_database():
    init_db()
    ensure_leave_tables()
    add_leave_payroll_columns()
    add_payment_order_id_column()


# Render / Gunicorn ke liye database setup
setup_database()


@app.route("/logout")
@login_required
def logout():
    session.clear()
    flash("Logged out successfully.", "success")
    return redirect(url_for("login"))
   

if __name__ == "__main__":
    app.run(debug=True)
