import os
import sqlite3
import datetime
import zipfile
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

RAZORPAY_KEY_ID = "rzp_test_SfKO3IFwsgnWhC"
RAZORPAY_SECRET = "F7CyUtorZhRe4q8YXtIefnv9"
razorpay_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_SECRET))

PAYMENTS_ENABLED = False
DEMO_MODE = True
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

    updated_at TEXT,
    FOREIGN KEY(company_id) REFERENCES companies(id)
)
""")

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

    return username in [admin_username.lower() for admin_username in ADMIN_USERNAMES]

def validate_required_columns(df, required_columns):
    df.columns = df.columns.str.strip()
    return [col for col in required_columns if col not in df.columns]


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

    conn = get_db()
    cur = conn.cursor()

    today = datetime.datetime.now().strftime("%Y-%m-%d")

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
          AND date(end_date) >= date(?)
        ORDER BY id DESC
        LIMIT 1
    """, (company_id, today))

    sub = cur.fetchone()
    conn.close()

    if sub:
        return {
            "is_pro": True,
            "plan": sub["plan_name"],
            "start_date": sub["start_date"],
            "end_date": sub["end_date"]
        }

    return {
        "is_pro": False,
        "plan": "FREE",
        "start_date": "-",
        "end_date": "-"
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
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    return get_active_plan()["is_pro"]


def require_pro_feature(message="Upgrade to PRO to use this feature."):
    # Admin/demo owner ko all PRO features allowed rahenge
    if is_admin_user():
        return True

    active_plan = get_active_plan()

    if active_plan and active_plan.get("plan", "").lower() != "free":
        return True

    flash(message, "warning")
    return False


def get_employee_limit():
    return None if is_pro_user() else 10


def can_add_employee():
    # Admin ko limit nahi lagegi
    if is_admin_user():
        return True, ""

    active_plan = get_active_plan()

    if active_plan and active_plan.get("plan", "").lower() != "free":
        return True, ""

    company_id = current_company_id()

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT COUNT(*) AS count
        FROM employees
        WHERE company_id = ?
    """, (company_id,))

    employee_count = cur.fetchone()["count"] or 0
    conn.close()

    if employee_count >= FREE_EMPLOYEE_LIMIT:
        return False, f"Free demo allows up to {FREE_EMPLOYEE_LIMIT} employees only. Paid subscription will be available soon."

    return True, ""


# ---------------------------
# COMPLIANCE RULES
# ---------------------------
def calculate_professional_tax_maharashtra(gross_salary, gender, payroll_month):
    gross_salary = float(gross_salary or 0)
    gender = str(gender or "male").strip().lower()
    month = month_only(payroll_month)

    if gender == "female":
        if gross_salary <= 25000:
            return 0
        return 300 if month == "02" else 200

    if gross_salary <= 7500:
        return 0

    if gross_salary <= 10000:
        return 175

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

    conn = get_db()
    cur = conn.cursor()

    if request.method == "POST":
        def to_float(name, default=0):
            try:
                return float(request.form.get(name, default) or default)
            except Exception:
                return default

        pf_employee_rate = to_float("pf_employee_rate", 12)
        pf_employer_rate = to_float("pf_employer_rate", 12)
        pf_wage_ceiling = to_float("pf_wage_ceiling", 15000)
        pf_max_deduction = to_float("pf_max_deduction", 1800)

        esic_employee_rate = to_float("esic_employee_rate", 0.75)
        esic_employer_rate = to_float("esic_employer_rate", 3.25)
        esic_wage_limit = to_float("esic_wage_limit", 21000)

        gratuity_rate = to_float("gratuity_rate", 4.81)
        bonus_rate = to_float("bonus_rate", 8.33)

        tds_enabled = 1 if request.form.get("tds_enabled") == "1" else 0

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
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            now
        ))

        conn.commit()
        conn.close()

        flash("Compliance settings updated successfully.")
        return redirect(url_for("compliance_settings"))

    conn.close()

    settings = get_compliance_settings(company_id)

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

    conn = get_db()
    cur = conn.cursor()

    if request.method == "POST":
        company_name = request.form.get("company_name", "").strip()
        address = request.form.get("address", "").strip()
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()

        try:
            overtime_multiplier = float(request.form.get("overtime_multiplier", 1) or 1)
        except Exception:
            overtime_multiplier = 1

        if overtime_multiplier not in [1, 2]:
            overtime_multiplier = 1

        working_days_policy = request.form.get("working_days_policy", "attendance").strip()

        if working_days_policy not in ["attendance", "fixed_26", "fixed_30"]:
            working_days_policy = "attendance"

        if not company_name:
            conn.close()
            flash("Company name is required.")
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
        conn.close()

        flash("Company profile updated successfully.")
        return redirect(url_for("company_profile"))

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
    """, (company_id,))

    company = cur.fetchone()
    conn.close()

    return render_template("company_profile.html", company=company)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


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

    cur.execute("""
        SELECT COUNT(*) AS count
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
    """, (company_id,))
    payroll_count = cur.fetchone()["count"] or 0

    cur.execute("""
        SELECT 
            COALESCE(SUM(gross), 0) AS total_gross,
            COALESCE(SUM(net_pay), 0) AS total_net_pay,
            COALESCE(SUM(total_deductions), 0) AS total_deductions
        FROM payroll_history
        WHERE company_id = ?
          AND is_current = 1
    """, (company_id,))
    payroll_totals = cur.fetchone()

    total_gross = round(float(payroll_totals["total_gross"] or 0))
    total_net_pay = round(float(payroll_totals["total_net_pay"] or 0))
    total_deductions = round(float(payroll_totals["total_deductions"] or 0))

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
        ORDER BY id DESC
        LIMIT 1
    """, (company_id,))
    subscription = cur.fetchone()

    cur.execute("""
        SELECT 
            month,
            COALESCE(SUM(gross), 0) AS gross,
            COALESCE(SUM(net_pay), 0) AS net_pay,
            COALESCE(SUM(total_deductions), 0) AS deductions
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

    for row in chart_rows:
        chart_labels.append(row["month"])
        chart_gross.append(round(float(row["gross"] or 0)))
        chart_net_pay.append(round(float(row["net_pay"] or 0)))
        chart_deductions.append(round(float(row["deductions"] or 0)))

    conn.close()

    active_plan = get_active_plan()

    if subscription:
        plan_name = subscription["plan_name"] or active_plan["plan"]
        subscription_status = subscription["status"] or "active"
        subscription_end_date = subscription["end_date"] or active_plan["end_date"] or "-"
    else:
        plan_name = active_plan["plan"]
        subscription_status = "free"
        subscription_end_date = active_plan["end_date"] or "-"

    return render_template(
        "dashboard.html",
        company_name=company["company_name"],

        employee_count=employee_count,
        attendance_count=attendance_count,
        payroll_count=payroll_count,
        pending_leaves=pending_leaves,

        total_gross=total_gross,
        total_net_pay=total_net_pay,
        total_deductions=total_deductions,

        chart_labels=chart_labels,
        chart_gross=chart_gross,
        chart_net_pay=chart_net_pay,
        chart_deductions=chart_deductions,

        active_plan=active_plan,
        plan_name=plan_name,
        subscription_status=subscription_status,
        subscription_end_date=subscription_end_date
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

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
        ORDER BY id DESC
        LIMIT 1
    """, (company_id,))

    active_subscription = cur.fetchone()
    conn.close()

    return render_template(
        "pricing.html",
        razorpay_key_id=RAZORPAY_KEY_ID,
        plans=plans,
        active_subscription=active_subscription,
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

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            id,
            amount,
            payment_id,
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
        elif status in ["pending", "created"]:
            pending_payments += 1
        else:
            failed_payments += 1

    total_paid_amount = round(total_paid_amount)

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ?
          AND status = 'active'
        ORDER BY id DESC
        LIMIT 1
    """, (company_id,))

    active_subscription = cur.fetchone()

    conn.close()

    return render_template(
        "payments.html",
        data=data,

        total_payments=total_payments,
        successful_payments=successful_payments,
        pending_payments=pending_payments,
        failed_payments=failed_payments,
        total_paid_amount=total_paid_amount,

        active_subscription=active_subscription
    )


# ---------------------------
# EMPLOYEES
# ---------------------------
@app.route("/upload-employees", methods=["GET", "POST"])
@login_required
def upload_employees():
    if request.method == "POST":
        company_id = current_company_id()

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

            missing_columns = validate_required_columns(df, required_columns)

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

            can_add, error_message = can_add_employee()

            if not can_add:
                flash(error_message, "warning")
                return redirect(url_for("pricing"))

            # IMPORTANT: conn/cur yahin create hona chahiye,
            # kyunki neeche employee limit aur insert queries me cur use hoga.
            conn = get_db()
            cur = conn.cursor()

            # Free demo upload limit check
            if not is_admin_user():
                active_plan = get_active_plan()

                if active_plan and active_plan.get("plan", "").lower() == "free":
                    cur.execute("""
                        SELECT COUNT(*) AS count
                        FROM employees
                        WHERE company_id = ?
                    """, (company_id,))

                    existing_count = cur.fetchone()["count"] or 0
                    upload_count = len(df)

                    if existing_count + upload_count > FREE_EMPLOYEE_LIMIT:
                        flash(
                            f"Free demo allows up to {FREE_EMPLOYEE_LIMIT} employees only. "
                            f"You already have {existing_count} employee(s), and your file contains {upload_count} row(s).",
                            "warning"
                        )
                        return redirect(url_for("upload_employees"))

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

                # Check whether employee already exists
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



    return render_template("employees.html", employees=employees)
@app.route("/employees")
@login_required
def employees_list():
    department = request.args.get("department", "").strip()
    search = request.args.get("search", "").strip()

    company_id = current_company_id()

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

    # Summary stats for Employee page
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

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Employee Master")

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
                "Must be unique for each employee.",
                "Employee full name.",
                "Designation or job role.",
                "Blank department will be treated as General.",
                "Use male, female, or other.",
                "Monthly salary must be greater than 0.",
                "Blank value will be treated as new.",
                "Use 0 if not applicable.",
                "Use 0 if not applicable.",
                "Keep as text to avoid number formatting issues.",
                "Keep as text to avoid number formatting issues.",
                "Employee bank name.",
                "Keep as text to avoid number formatting issues.",
                "Use uppercase IFSC code."
            ]
        })

        instructions.to_excel(writer, index=False, sheet_name="Instructions")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
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

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")

            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(cell_value))

                ws.column_dimensions[column_letter].width = max_length + 4

        # Keep UAN, ESIC, Account No and IFSC as text in Employee Master sheet
        ws = workbook["Employee Master"]
        text_columns = ["J", "K", "M", "N"]

        for col in text_columns:
            for cell in ws[col]:
                cell.number_format = "@"

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
                "overtime_hours",
                "bonus",
                "manual_deduction"
            ]

            missing_columns = validate_required_columns(df, required_columns)

            if missing_columns:
                session["error_report"] = create_error_report(
                    [f"Missing required column: {col}" for col in missing_columns],
                    "attendance_upload_errors.xlsx"
                )
                flash("Upload failed. Required columns are missing. Please download the error report.", "danger")
                return redirect(url_for("upload_attendance"))

            row_errors = []

            # Clean helper columns
            df["emp_code_clean"] = df["emp_code"].apply(lambda x: clean_text(x))
            df["month_clean"] = df["month"].apply(lambda x: clean_text(x))

            # Duplicate check: same employee + same month inside uploaded file
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
                overtime_hours = clean_float(row.get("overtime_hours"), -1)
                bonus = clean_float(row.get("bonus"), -1)
                manual_deduction = clean_float(row.get("manual_deduction"), -1)

                if emp_code == "":
                    row_errors.append(f"Row {row_no}: emp_code missing")

                if month == "":
                    row_errors.append(f"Row {row_no}: month missing")
                else:
                    # Required month format: YYYY-MM
                    try:
                        datetime.datetime.strptime(month, "%Y-%m")
                    except Exception:
                        row_errors.append(f"Row {row_no}: month must be in YYYY-MM format, example 2026-12")

                if working_days <= 0:
                    row_errors.append(f"Row {row_no}: working_days must be greater than 0")

                if present_days < 0:
                    row_errors.append(f"Row {row_no}: present_days cannot be negative")

                if overtime_hours < 0:
                    row_errors.append(f"Row {row_no}: overtime_hours cannot be negative")

                if bonus < 0:
                    row_errors.append(f"Row {row_no}: bonus cannot be negative")

                if manual_deduction < 0:
                    row_errors.append(f"Row {row_no}: manual_deduction cannot be negative")

                if working_days > 0 and present_days > working_days:
                    row_errors.append(f"Row {row_no}: present_days cannot be greater than working_days")

                if working_days > 31:
                    row_errors.append(f"Row {row_no}: working_days cannot be greater than 31")

                if present_days > 31:
                    row_errors.append(f"Row {row_no}: present_days cannot be greater than 31")

            if row_errors:
                session["error_report"] = create_error_report(
                    row_errors,
                    "attendance_upload_errors.xlsx"
                )
                flash("Upload failed. Please download the error report and fix the file.", "danger")
                return redirect(url_for("upload_attendance"))

            conn = get_db()
            cur = conn.cursor()

            # Check employee codes exist in Employee Master
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

            # Delete existing attendance only for uploaded months
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

                working_days = int(clean_float(row.get("working_days"), 0))
                present_days = int(clean_float(row.get("present_days"), 0))
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
                        overtime_hours,
                        bonus,
                        manual_deduction
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    company_id,
                    emp_code,
                    month,
                    working_days,
                    present_days,
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

    conn = get_db()
    cur = conn.cursor()

    query = """
        SELECT 
            a.*,
            e.employee_name,
            e.role,
            e.department
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

    query += " ORDER BY a.id DESC"

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

    total_records = len(rows)
    total_working_days = round(sum(float(r["working_days"] or 0) for r in rows))
    total_present_days = round(sum(float(r["present_days"] or 0) for r in rows))
    total_absent_days = round(total_working_days - total_present_days)

    if total_absent_days < 0:
        total_absent_days = 0

    total_overtime_hours = round(sum(float(r["overtime_hours"] or 0) for r in rows), 2)
    total_bonus = round(sum(float(r["bonus"] or 0) for r in rows))
    total_manual_deduction = round(sum(float(r["manual_deduction"] or 0) for r in rows))

    if total_working_days > 0:
        attendance_percentage = round((total_present_days / total_working_days) * 100, 2)
    else:
        attendance_percentage = 0

    conn.close()

    return render_template(
        "attendance.html",
        rows=rows,
        departments=departments,

        selected_month=month,
        selected_department=department,
        search=search,

        total_records=total_records,
        total_working_days=total_working_days,
        total_present_days=total_present_days,
        total_absent_days=total_absent_days,
        total_overtime_hours=total_overtime_hours,
        total_bonus=total_bonus,
        total_manual_deduction=total_manual_deduction,
        attendance_percentage=attendance_percentage
    )


@app.route("/leave-management", methods=["GET", "POST"])
@login_required
def leave_management():
    company_id = current_company_id()
    conn = get_db()
    cur = conn.cursor()

    if request.method == "POST":
        emp_code = request.form["emp_code"].strip()
        leave_type = request.form["leave_type"].strip()
        start_date = request.form["start_date"]
        end_date = request.form["end_date"]
        total_days = float(request.form.get("total_days") or 0)
        reason = request.form.get("reason", "").strip()

        cur.execute("""
            INSERT INTO leave_requests
            (company_id, emp_code, leave_type, start_date, end_date, total_days, reason, status)
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
        conn.close()

        flash("Leave request added successfully.")
        return redirect(url_for("leave_management"))

    cur.execute("""
        INSERT OR IGNORE INTO leave_policy_settings
        (company_id, casual_leave_limit, sick_leave_limit, paid_leave_limit)
        VALUES (?, 6, 6, 12)
    """, (company_id,))

    cur.execute("""
        SELECT casual_leave_limit,
               sick_leave_limit,
               paid_leave_limit
        FROM leave_policy_settings
        WHERE company_id = ?
    """, (company_id,))
    leave_policy = cur.fetchone()

    cur.execute("""
        SELECT emp_code, employee_name, department
        FROM employees
        WHERE company_id = ?
        ORDER BY employee_name
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
            leave_policy["casual_leave_limit"],
            leave_policy["sick_leave_limit"],
            leave_policy["paid_leave_limit"]
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

    conn.close()

    return render_template(
        "leave_management.html",
        employees=employees,
        leave_requests=leave_requests,
        leave_balances=leave_balances,
        leave_policy=leave_policy
    )

@app.route("/update-leave-policy", methods=["POST"])
@login_required
def update_leave_policy():
    company_id = current_company_id()

    casual_leave_limit = float(request.form.get("casual_leave_limit") or 0)
    sick_leave_limit = float(request.form.get("sick_leave_limit") or 0)
    paid_leave_limit = float(request.form.get("paid_leave_limit") or 0)

    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO leave_policy_settings
        (company_id, casual_leave_limit, sick_leave_limit, paid_leave_limit)
        VALUES (?, ?, ?, ?)
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

    conn.commit()
    conn.close()

    flash("Leave policy updated successfully.")
    return redirect(url_for("leave_management"))


@app.route("/approve-leave/<int:leave_id>", methods=["POST"])
@login_required
def approve_leave(leave_id):
    company_id = current_company_id()
    conn = get_db()
    cur = conn.cursor()

    try:
        # 1. Leave request fetch
        cur.execute("""
            SELECT *
            FROM leave_requests
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        leave = cur.fetchone()

        if not leave:
            flash("Leave request not found.", "danger")
            return redirect(url_for("leave_management"))

        # 2. Double approval avoid
        if leave["status"] == "Approved":
            flash("This leave is already approved.", "warning")
            return redirect(url_for("leave_management"))

        if leave["status"] == "Rejected":
            flash("Rejected leave cannot be approved directly. Please create a new leave request.", "warning")
            return redirect(url_for("leave_management"))

        emp_code = leave["emp_code"]
        leave_type = leave["leave_type"]
        total_days = float(leave["total_days"] or 0)

        # 3. Leave type ko balance table column se map karo
        leave_column_map = {
            "Casual Leave": "casual_leave",
            "Sick Leave": "sick_leave",
            "Paid Leave": "paid_leave"
        }

        # Unpaid Leave ka balance minus nahi hoga
        if leave_type == "Unpaid Leave":
            cur.execute("""
                UPDATE leave_requests
                SET status = 'Approved'
                WHERE id = ?
                  AND company_id = ?
            """, (leave_id, company_id))

            conn.commit()
            flash("Unpaid leave approved successfully. Leave balance not changed.", "success")
            return redirect(url_for("leave_management"))

        if leave_type not in leave_column_map:
            flash(f"Invalid leave type: {leave_type}", "danger")
            return redirect(url_for("leave_management"))

        balance_column = leave_column_map[leave_type]

        # 4. Employee leave balance fetch
        cur.execute("""
            SELECT *
            FROM leave_balances
            WHERE company_id = ?
              AND emp_code = ?
        """, (company_id, emp_code))

        balance_row = cur.fetchone()

        if not balance_row:
            flash("Leave balance not found for this employee. Please create leave balance first.", "danger")
            return redirect(url_for("leave_management"))

        current_balance = float(balance_row[balance_column] or 0)
        used_leave = float(balance_row["used_leave"] or 0)

        # 5. Balance check
        if current_balance < total_days:
            flash(
                f"Insufficient {leave_type} balance. Available: {current_balance}, Required: {total_days}",
                "danger"
            )
            return redirect(url_for("leave_management"))

        new_balance = current_balance - total_days
        new_used_leave = used_leave + total_days

        # 6. Balance minus + used leave increase
        cur.execute(f"""
            UPDATE leave_balances
            SET {balance_column} = ?,
                used_leave = ?
            WHERE company_id = ?
              AND emp_code = ?
        """, (new_balance, new_used_leave, company_id, emp_code))

        # 7. Leave request approve
        cur.execute("""
            UPDATE leave_requests
            SET status = 'Approved'
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        conn.commit()

        flash(
            f"Leave approved successfully. {leave_type} balance updated: {current_balance} → {new_balance}",
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
    conn = get_db()
    cur = conn.cursor()

    try:
        # 1. Leave request fetch
        cur.execute("""
            SELECT *
            FROM leave_requests
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        leave = cur.fetchone()

        if not leave:
            flash("Leave request not found.", "danger")
            return redirect(url_for("leave_management"))

        # 2. Approved leave direct reject nahi karenge
        if leave["status"] == "Approved":
            flash("Approved leave cannot be rejected directly. Cancel/reversal logic is required.", "warning")
            return redirect(url_for("leave_management"))

        # 3. Reject only status change karega, balance nahi
        cur.execute("""
            UPDATE leave_requests
            SET status = 'Rejected'
            WHERE id = ?
              AND company_id = ?
        """, (leave_id, company_id))

        conn.commit()
        flash("Leave rejected successfully. Leave balance not changed.", "success")

    except Exception as e:
        conn.rollback()
        flash(f"Error while rejecting leave: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("leave_management"))


@app.route("/download-attendance-sample")
@login_required
def download_attendance_sample():
    data = {
        "emp_code": ["EMP001", "EMP002"],
        "month": ["2026-12", "2026-12"],
        "working_days": [30, 30],
        "present_days": [26, 24],
        "overtime_hours": [2, 2],
        "bonus": [0, 0],
        "manual_deduction": [0, 0]
    }

    df = pd.DataFrame(data)

    file_path = os.path.join(UPLOAD_FOLDER, "attendance_sample.xlsx")

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Attendance")

        instructions = pd.DataFrame({
            "Field": [
                "emp_code",
                "month",
                "working_days",
                "present_days",
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
                "Yes"
            ],
            "Example": [
                "EMP001",
                "2026-12",
                "30",
                "26",
                "2",
                "0",
                "0"
            ],
            "Notes": [
                "Employee code must exist in Employee Master.",
                "Use YYYY-MM format only, example 2026-12.",
                "Total working days for selected month or company policy.",
                "Present days cannot be greater than working days.",
                "Use 0 if no overtime.",
                "Use 0 if no attendance bonus.",
                "Use 0 if no manual deduction."
            ]
        })

        instructions.to_excel(writer, index=False, sheet_name="Instructions")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
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

            for row in ws.iter_rows(min_row=2):
                for cell in row:
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
        ws = workbook["Attendance"]

        for col in ["A", "B"]:
            for cell in ws[col]:
                cell.number_format = "@"

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

    # Company policies: overtime + working days
    cur.execute("""
        SELECT 
            COALESCE(overtime_multiplier, 1) AS overtime_multiplier,
            COALESCE(working_days_policy, 'attendance') AS working_days_policy
        FROM companies
        WHERE id = ?
    """, (company_id,))
    company = cur.fetchone()

    overtime_multiplier = 1
    working_days_policy = "attendance"

    if company:
        try:
            overtime_multiplier = float(company["overtime_multiplier"] or 1)
        except Exception:
            overtime_multiplier = 1

        working_days_policy = company["working_days_policy"] or "attendance"

    if overtime_multiplier not in [1, 2]:
        overtime_multiplier = 1

    if working_days_policy not in ["attendance", "fixed_26", "fixed_30"]:
        working_days_policy = "attendance"

    # Compliance settings
    settings = get_compliance_settings(company_id)

    pf_employee_rate = float(settings["pf_employee_rate"] or 12) / 100
    pf_employer_rate = float(settings["pf_employer_rate"] or 12) / 100
    pf_max_deduction = float(settings["pf_max_deduction"] or 1800)

    esic_employee_rate = float(settings["esic_employee_rate"] or 0.75) / 100
    esic_employer_rate = float(settings["esic_employer_rate"] or 3.25) / 100
    esic_wage_limit = float(settings["esic_wage_limit"] or 21000)

    gratuity_rate = float(settings["gratuity_rate"] or 4.81) / 100
    bonus_rate = float(settings["bonus_rate"] or 8.33) / 100

    # Mark old payroll of same month as not current
    cur.execute("""
        UPDATE payroll_history
        SET is_current = 0
        WHERE company_id = ?
          AND month = ?
    """, (company_id, month))

    cur.execute("""
        SELECT 
            e.*,
            a.working_days,
            a.present_days,
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

    for row in rows:
        emp_code = row["emp_code"]
        monthly_salary = float(row["monthly_salary"] or 0)

        attendance_working_days = int(row["working_days"] or 30)

        if working_days_policy == "fixed_26":
            working_days = 26
        elif working_days_policy == "fixed_30":
            working_days = 30
        else:
            working_days = attendance_working_days

        if working_days <= 0:
            working_days = 30

        present_days = float(row["present_days"] or 0)
        overtime_hours = float(row["overtime_hours"] or 0)
        manual_deduction = float(row["manual_deduction"] or 0)
        attendance_bonus = float(row["bonus"] or 0)

        gender = str(row["gender"] or "male").strip().lower()

        # Approved leave calculation for selected month
        cur.execute("""
            SELECT
                COALESCE(SUM(
                    CASE 
                        WHEN leave_type IN ('Casual Leave', 'Sick Leave', 'Paid Leave')
                        THEN total_days
                        ELSE 0
                    END
                ), 0) AS paid_leave_days,

                COALESCE(SUM(
                    CASE 
                        WHEN leave_type = 'Unpaid Leave'
                        THEN total_days
                        ELSE 0
                    END
                ), 0) AS lwp_days
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

        paid_leave_days = float(leave_data["paid_leave_days"] or 0)
        lwp_days = float(leave_data["lwp_days"] or 0)

        per_day_salary = monthly_salary / working_days

        # Paid leave salary protect karega
        payable_days = present_days + paid_leave_days

        # Safety: payable days working days se zyada nahi hone chahiye
        if payable_days > working_days:
            payable_days = working_days

        # LWP display amount
        lwp_deduction = rupee(per_day_salary * lwp_days)

        # Salary payable days ke basis par calculate hogi
        earned_salary = per_day_salary * payable_days

        basic = rupee(earned_salary * 0.40)
        da = rupee(earned_salary * 0.10)
        hra = rupee(earned_salary * 0.20)

        special_allowance = float(row["special_allowance"] or 0)

        # Agar employee ne full month kaam nahi kiya hai, special allowance bhi proportionate karo
        special_allowance = rupee((special_allowance / working_days) * payable_days)

        other_allowance = earned_salary - basic - da - hra - special_allowance

        if other_allowance < 0:
            other_allowance = 0

        other_allowance = rupee(other_allowance)

        gross = rupee(
            basic
            + da
            + hra
            + special_allowance
            + other_allowance
        )

        # Overtime calculation with company-wise multiplier
        if overtime_hours > 0:
            hourly_rate = monthly_salary / 30 / 8
            overtime_amount = rupee(
                hourly_rate
                * overtime_hours
                * overtime_multiplier
            )
        else:
            overtime_amount = 0

        # PF calculation
        pf_base = basic + da

        pf_employee = min(
            rupee(pf_base * pf_employee_rate),
            pf_max_deduction
        )

        pf_employer = min(
            rupee(pf_base * pf_employer_rate),
            pf_max_deduction
        )

        # ESIC calculation
        if gross <= esic_wage_limit:
            esi_employee = round(gross * esic_employee_rate)
            esi_employer = round(gross * esic_employer_rate)
        else:
            esi_employee = 0
            esi_employer = 0

        professional_tax = rupee(
            calculate_professional_tax_maharashtra(gross, gender, month)
        )

        lwf = calculate_lwf_maharashtra(month)
        lwf_employee = rupee(lwf["employee"])
        lwf_employer = rupee(lwf["employer"])

        tds = 0

        bonus_ctc, festival_bonus = calculate_bonus_logic(
            basic,
            month,
            bonus_rate
        )

        festival_bonus = rupee(festival_bonus + attendance_bonus)

        # Important:
        # lwp_deduction ko total_deductions me add nahi kar rahe,
        # kyunki gross already payable_days ke basis par reduced hai.
        # Agar yaha bhi add karenge to double deduction ho jayega.
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

    conn = get_db()
    cur = conn.cursor()

    query = """
        SELECT 
            *,
            COALESCE(payable_days, 0) AS payable_days,
            COALESCE(paid_leave_days, 0) AS paid_leave_days,
            COALESCE(lwp_days, 0) AS lwp_days,
            COALESCE(lwp_deduction, 0) AS lwp_deduction
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
            )
        """
        search_value = f"%{search}%"
        params.extend([search_value, search_value, search_value])

    query += " ORDER BY id DESC"

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
    total_net_pay = round(sum(float(row["net_pay"] or 0) for row in records))
    total_deductions = round(sum(float(row["total_deductions"] or 0) for row in records))

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
        total_net_pay=total_net_pay,
        total_deductions=total_deductions,

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
        flash("Please select month to export payroll")
        return redirect(url_for("payroll_history"))

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

            COALESCE(p.payable_days, 0) AS payable_days,
            COALESCE(p.paid_leave_days, 0) AS paid_leave_days,
            COALESCE(p.lwp_days, 0) AS lwp_days,
            COALESCE(p.lwp_deduction, 0) AS lwp_deduction,

            p.basic,
            p.da,
            p.hra,
            p.special_allowance,
            p.other_allowance,
            p.gross,

            p.esi_employee,
            p.professional_tax,
            p.pf_employee,
            p.lwf_employee,
            p.tds,
            p.manual_deduction,
            p.total_deductions,

            p.esi_employer,
            p.pf_employer,
            p.gratuity,
            p.bonus_ctc,
            p.festival_bonus,
            p.lwf_employer,
            p.total_contributions,

            p.net_pay,
            p.monthly_ctc,
            p.annual_ctc,

            p.overtime_hours,
            p.overtime_amount,
            p.created_at

        FROM payroll_history p

        LEFT JOIN employees e
          ON p.company_id = e.company_id
         AND p.emp_code = e.emp_code

        WHERE p.company_id = ?
          AND p.month = ?
          AND p.is_current = 1
    """

    params = [current_company_id(), month]

    if department:
        query += " AND p.department = ?"
        params.append(department)

    query += " ORDER BY p.id DESC"

    df = pd.read_sql_query(query, conn, params=tuple(params))
    conn.close()

    if df.empty:
        flash("No payroll data found for selected month")
        return redirect(url_for("payroll_history"))

    money_columns = [
        "monthly_salary",

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
        "bonus_ctc",
        "festival_bonus",
        "lwf_employer",
        "total_contributions",
        "net_pay",
        "monthly_ctc",
        "annual_ctc",
        "overtime_hours",
        "overtime_amount"
    ]

    for col in money_columns:
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

    # Professional column names for Excel
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

        "payable_days": "Payable Days",
        "paid_leave_days": "Paid Leave Days",
        "lwp_days": "LWP Days",
        "lwp_deduction": "LWP Deduction",

        "basic": "Basic",
        "da": "DA",
        "hra": "HRA",
        "special_allowance": "Special Allowance",
        "other_allowance": "Other Allowance",
        "gross": "Gross Salary",

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
        "bonus_ctc": "Bonus CTC",
        "festival_bonus": "Festival Bonus",
        "lwf_employer": "LWF Employer",
        "total_contributions": "Total Contributions",

        "net_pay": "Net Pay",
        "monthly_ctc": "Monthly CTC",
        "annual_ctc": "Annual CTC",

        "overtime_hours": "Overtime Hours",
        "overtime_amount": "Overtime Amount",
        "created_at": "Created At"
    })

    file_name = f"payroll_{month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Payroll Register")

        ws = writer.book["Payroll Register"]

        # Freeze header row
        ws.freeze_panes = "A2"

        # Enable filter
        ws.auto_filter.ref = ws.dimensions

        # Header formatting
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Auto column width
        for column_cells in ws.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                cell_value = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(cell_value))

            ws.column_dimensions[column_letter].width = max_length + 3

        # Keep UAN, ESIC, Account No, IFSC as text
        text_format_headers = ["UAN No", "ESIC No", "Account No", "IFSC Code"]

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
            a.overtime_hours,
            a.bonus,
            a.manual_deduction
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

            gross,
            pf_employee,
            esi_employee,
            professional_tax,
            total_deductions,
            net_pay,
            monthly_ctc,
            annual_ctc,
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

    # Employee Master Audit flags
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

    # Month-wise attendance/payroll check
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

    # Leave summary numbers
    if payroll_df.empty:
        total_paid_leave_days = 0
        total_lwp_days = 0
        total_lwp_deduction = 0
        employees_with_lwp = 0
    else:
        total_paid_leave_days = round(float(payroll_df["paid_leave_days"].fillna(0).sum()), 2)
        total_lwp_days = round(float(payroll_df["lwp_days"].fillna(0).sum()), 2)
        total_lwp_deduction = round(float(payroll_df["lwp_deduction"].fillna(0).sum()))
        employees_with_lwp = int((payroll_df["lwp_days"].fillna(0) > 0).sum())

    approved_leave_count = 0
    rejected_leave_count = 0
    pending_leave_count = 0

    if not leave_df.empty:
        approved_leave_count = int((leave_df["status"] == "Approved").sum())
        rejected_leave_count = int((leave_df["status"] == "Rejected").sum())
        pending_leave_count = int((leave_df["status"] == "Pending").sum())

    summary_data = {
        "Audit Month": [month] * 20,
        "Audit Item": [
            "Total Employees",
            "Attendance Uploaded",
            "Payroll Processed",

            "Total Paid Leave Days",
            "Total LWP Days",
            "Total LWP Deduction",
            "Employees With LWP",
            "Approved Leave Requests",
            "Rejected Leave Requests",
            "Pending Leave Requests",

            "Missing UAN",
            "Missing ESIC No.",
            "Missing Bank Name",
            "Missing Account No.",
            "Missing IFSC",
            "Missing Department",
            "Missing Gender",
            "Invalid Salary",
            "Attendance Missing",
            "Payroll Missing"
        ],
        "Count": [
            len(employees_df),
            len(attendance_emp_codes),
            len(payroll_emp_codes),

            total_paid_leave_days,
            total_lwp_days,
            total_lwp_deduction,
            employees_with_lwp,
            approved_leave_count,
            rejected_leave_count,
            pending_leave_count,

            int((employees_df["uan_status"] == "Missing").sum()),
            int((employees_df["esic_status"] == "Missing").sum()),
            int((employees_df["bank_status"] == "Missing").sum()),
            int((employees_df["account_status"] == "Missing").sum()),
            int((employees_df["ifsc_status"] == "Missing").sum()),
            int((employees_df["department_status"] == "Missing").sum()),
            int((employees_df["gender_status"] == "Missing").sum()),
            int((employees_df["salary_status"] == "Invalid").sum()),
            int((employees_df["attendance_status"] == "Attendance Missing").sum()),
            int((employees_df["payroll_status"] == "Payroll Missing").sum())
        ]
    }

    summary_df = pd.DataFrame(summary_data)

    payroll_missing_df = employee_audit_df[
        employee_audit_df["payroll_status"] == "Payroll Missing"
    ]

    attendance_missing_df = employee_audit_df[
        employee_audit_df["attendance_status"] == "Attendance Missing"
    ]

    file_name = f"hr_audit_report_{month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Audit Summary")
        employee_audit_df.to_excel(writer, index=False, sheet_name="Employee Master Audit")
        attendance_df.to_excel(writer, index=False, sheet_name="Attendance Audit")
        payroll_df.to_excel(writer, index=False, sheet_name="Payroll Audit")
        leave_df.to_excel(writer, index=False, sheet_name="Leave Audit")
        attendance_missing_df.to_excel(writer, index=False, sheet_name="Attendance Missing")
        payroll_missing_df.to_excel(writer, index=False, sheet_name="Payroll Missing")

        workbook = writer.book

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        missing_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        ok_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

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

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_name
    )


@app.route("/full-and-final", methods=["GET", "POST"])
@login_required
def full_and_final():
    company_id = current_company_id()
    conn = get_db()
    cur = conn.cursor()

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

    if request.method == "POST":
        try:
            emp_code = request.form.get("emp_code", "").strip()
            last_working_day = request.form.get("last_working_day", "").strip()
            settlement_month = request.form.get("settlement_month", "").strip()

            reason = request.form.get("reason", "").strip()
            remarks = request.form.get("remarks", "").strip()

            # Optional manual adjustment fields.
            # HR can keep these blank/0. System will still calculate F&F automatically.
            bonus_payable = to_float(request.form.get("bonus_payable"), 0)
            gratuity_payable = to_float(request.form.get("gratuity_payable"), 0)
            other_earnings = to_float(request.form.get("other_earnings"), 0)

            notice_recovery = to_float(request.form.get("notice_recovery"), 0)
            loan_recovery = to_float(request.form.get("loan_recovery"), 0)
            advance_recovery = to_float(request.form.get("advance_recovery"), 0)
            other_deductions = to_float(request.form.get("other_deductions"), 0)

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
                    last_working_date_obj = datetime.datetime.strptime(last_working_day, "%Y-%m-%d")
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

            if errors:
                flash(" ".join(errors), "danger")
                return redirect(url_for("full_and_final"))

            cur.execute("""
                SELECT emp_code, employee_name, role, department, monthly_salary
                FROM employees
                WHERE company_id = ?
                  AND emp_code = ?
            """, (company_id, emp_code))

            emp = cur.fetchone()

            if not emp:
                flash("Employee not found.", "danger")
                return redirect(url_for("full_and_final"))

            monthly_salary = float(emp["monthly_salary"] or 0)

            if monthly_salary <= 0:
                flash("Employee monthly salary is invalid. Please update employee master.", "danger")
                return redirect(url_for("full_and_final"))

            # Auto paid days from last working day.
            # Example: Last Working Day 2026-12-06 => paid_days = 6
            paid_days = min(last_working_date_obj.day, 30)

            if paid_days < 0:
               paid_days = 0

            if paid_days > 31:
               paid_days = 31

            per_day_salary = monthly_salary / 30
            earned_salary = money_round(per_day_salary * paid_days)

            # Auto leave balance.
            # Professional rule: only Paid Leave is encashed.
            cur.execute("""
                SELECT 
                    COALESCE(paid_leave, 0) AS paid_leave
                FROM leave_balances
                WHERE company_id = ?
                  AND emp_code = ?
                LIMIT 1
            """, (company_id, emp_code))

            leave_row = cur.fetchone()

            if leave_row:
                leave_balance = float(leave_row["paid_leave"] or 0)
            else:
                leave_balance = 0

            apply_leave_encashment = request.form.get("apply_leave_encashment", "no")

            if apply_leave_encashment == "yes":
                leave_encashment = money_round(per_day_salary * leave_balance)
            else:
                leave_encashment = 0

            total_earnings = money_round(
                earned_salary
                + leave_encashment
                + bonus_payable
                + gratuity_payable
                + other_earnings
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
                f"Full & Final settlement created successfully for {emp['employee_name']}. Final Payable: ₹{final_payable}",
                "success"
            )

            return redirect(url_for("full_and_final"))

        except Exception as e:
            conn.rollback()
            flash(f"Error while creating Full & Final settlement: {str(e)}", "danger")
            return redirect(url_for("full_and_final"))

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
    total_final_payable = round(sum(float(row["final_payable"] or 0) for row in settlements))
    total_earnings = round(sum(float(row["total_earnings"] or 0) for row in settlements))
    total_deductions = round(sum(float(row["total_deductions"] or 0) for row in settlements))

    conn.close()

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

    emp_code = str(row["emp_code"] or "employee").replace(" ", "_")
    settlement_month = str(row["settlement_month"] or "month").replace(" ", "_")

    file_name = f"fnf_settlement_{emp_code}_{settlement_month}.xlsx"
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    data = [
        ["FULL & FINAL SETTLEMENT", ""],
        ["", ""],

        ["Company Details", ""],
        ["Company Name", row["company_name"] or "-"],
        ["Company Address", row["company_address"] or "-"],
        ["Company Email", row["company_email"] or "-"],
        ["Company Phone", row["company_phone"] or "-"],
        ["", ""],

        ["Employee Details", ""],
        ["Employee Code", row["emp_code"]],
        ["Employee Name", row["employee_name"]],
        ["Designation", row["role"]],
        ["Department", row["department"]],
        ["Last Working Day", row["last_working_day"]],
        ["Settlement Month", row["settlement_month"]],
        ["Reason for Leaving", row["reason"] or "-"],
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
        ["Remarks", row["remarks"] or "-"],
        ["Created At", row["created_at"]],
        ["", ""],

        ["Approvals", ""],
        ["Prepared By", ""],
        ["Checked By", ""],
        ["HR / Authorized Signatory", ""],
    ]

    df = pd.DataFrame(data, columns=["Particulars", "Details"])

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="F&F Settlement")

        ws = writer.book["F&F Settlement"]

        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        section_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        final_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
        warning_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

        header_font = Font(bold=True, color="FFFFFF")
        section_font = Font(bold=True, color="1E3A8A")
        final_font = Font(bold=True, color="166534")
        normal_font = Font(color="0F172A")

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0")
        )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Header row
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        # Main title row
        ws["A2"].font = Font(bold=True, size=16, color="0F172A")
        ws["A2"].alignment = Alignment(horizontal="center")
        ws.merge_cells("A2:B2")

        # Style all rows
        for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
            label = row_cells[0].value

            for cell in row_cells:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                cell.font = normal_font

            if label in [
                "Company Details",
                "Employee Details",
                "Earnings",
                "Deductions",
                "Final Settlement",
                "Approvals"
            ]:
                for cell in row_cells:
                    cell.font = section_font
                    cell.fill = section_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            if label in ["Total Earnings", "Total Deductions"]:
                for cell in row_cells:
                    cell.font = Font(bold=True, color="0F172A")
                    cell.fill = warning_fill

            if label == "Final Payable":
                for cell in row_cells:
                    cell.font = final_font
                    cell.fill = final_fill

            if label in ["Prepared By", "Checked By", "HR / Authorized Signatory"]:
                row_cells[1].value = "________________________"

        # Amount formatting
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

        for row_idx in range(1, ws.max_row + 1):
            label = ws.cell(row=row_idx, column=1).value
            if label in amount_labels:
                ws.cell(row=row_idx, column=2).number_format = '₹#,##0'

        # Auto width
        for column_cells in ws.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                cell_value = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(cell_value))

            ws.column_dimensions[column_letter].width = max_length + 5

        ws.row_dimensions[2].height = 28

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

    file_name = f"fnf_settlement_{row['emp_code']}_{row['settlement_month']}.pdf"
    file_path = os.path.join(PAYSLIP_FOLDER, file_name)

    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter

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

    def clean(value, default="-"):
        if value is None:
            return default

        value = str(value).strip()

        if value == "" or value.lower() in ["nan", "none", "null"]:
            return default

        return value

    def row_value(key, default=""):
        try:
            return row[key]
        except Exception:
            return default

    company_name = clean(row_value("company_name"), "SMART HIRE AI PAYROLL")
    company_address = clean(row_value("company_address"), "")
    company_email = clean(row_value("company_email"), "")
    company_phone = clean(row_value("company_phone"), "")

    company_contact_parts = []

    if company_address not in ["", "-"]:
        company_contact_parts.append(company_address)

    if company_email not in ["", "-"]:
        company_contact_parts.append(company_email)

    if company_phone not in ["", "-"]:
        company_contact_parts.append(company_phone)

    company_contact_line = " | ".join(company_contact_parts) if company_contact_parts else "-"

    y = height - 35

    # Header
    c.setFillColorRGB(0.10, 0.17, 0.28)
    c.rect(35, y - 58, width - 70, 58, fill=1, stroke=0)

    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 17)
    c.drawCentredString(width / 2, y - 21, company_name.upper())

    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, y - 39, "Full & Final Settlement Statement")
    c.drawCentredString(width / 2, y - 52, company_contact_line)

    y -= 88

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(width / 2, y, f"FULL & FINAL SETTLEMENT - {clean(row_value('settlement_month'))}")

    y -= 28

    # Employee details
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Employee & Exit Details")

    y -= 10
    c.rect(40, y - 118, width - 80, 118)

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
        ("Reason", row_value("reason")),
        ("Paid Days", number_text(row_value("paid_days", 0))),
        ("Created At", row_value("created_at")),
    ]

    yy = y - 22
    for label, value in details_left:
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(55, yy, f"{label}:")
        c.setFont("Helvetica", 8.5)
        c.drawString(155, yy, clean(value))
        yy -= 20

    yy = y - 22
    for label, value in details_right:
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(330, yy, f"{label}:")
        c.setFont("Helvetica", 8.5)
        c.drawString(435, yy, clean(value))
        yy -= 20

    y -= 150

    # Earnings and deductions table
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Settlement Calculation")
    y -= 15

    x = 40
    table_w = width - 80
    row_h = 22
    table_y = y
    rows_count = 10

    c.rect(x, table_y - row_h * rows_count, table_w, row_h * rows_count)

    c.setFillColorRGB(0.12, 0.36, 0.85)
    c.rect(x, table_y - row_h, table_w, row_h, fill=1, stroke=0)

    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x + 8, table_y - 14, "Earnings")
    c.drawString(x + table_w * 0.35 + 8, table_y - 14, "Amount")
    c.drawString(x + table_w * 0.50 + 8, table_y - 14, "Deductions")
    c.drawString(x + table_w * 0.78 + 8, table_y - 14, "Amount")

    c.setFillColorRGB(0, 0, 0)
    c.line(x + table_w * 0.35, table_y, x + table_w * 0.35, table_y - row_h * rows_count)
    c.line(x + table_w * 0.50, table_y, x + table_w * 0.50, table_y - row_h * rows_count)
    c.line(x + table_w * 0.78, table_y, x + table_w * 0.78, table_y - row_h * rows_count)

    for i in range(rows_count + 1):
        c.line(x, table_y - row_h * i, x + table_w, table_y - row_h * i)

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

    c.setFont("Helvetica", 8.5)
    start_y = table_y - row_h - 14

    for i in range(9):
        yy = start_y - row_h * i

        earning_label, earning_value, earning_type = earnings[i]

        c.drawString(x + 8, yy, earning_label)

        if earning_type == "number":
            c.drawRightString(x + table_w * 0.50 - 8, yy, number_text(earning_value))
        else:
            c.drawRightString(x + table_w * 0.50 - 8, yy, money(earning_value))

        deduction_label, deduction_value = deductions[i]

        if deduction_label:
            c.drawString(x + table_w * 0.50 + 8, yy, deduction_label)
            c.drawRightString(x + table_w - 8, yy, money(deduction_value))

    y = table_y - row_h * rows_count - 34

    # Final payable highlight
    c.setFillColorRGB(0.86, 0.96, 0.89)
    c.rect(40, y - 38, width - 80, 38, fill=1, stroke=0)

    c.setFillColorRGB(0, 0.35, 0.15)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(55, y - 24, f"FINAL PAYABLE: {money(row_value('final_payable', 0))}")

    y -= 62

    # Remarks box
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y, "Remarks")

    y -= 10
    c.rect(40, y - 38, width - 80, 38)

    c.setFont("Helvetica", 8.5)
    remarks = clean(row_value("remarks"), "-")

    if len(remarks) > 110:
        remarks = remarks[:107] + "..."

    c.drawString(55, y - 22, remarks)

    y -= 68

    # Signature section
    c.setFont("Helvetica", 8.5)
    c.drawString(40, y, "Prepared By")
    c.drawString(260, y, "Checked By")
    c.drawString(470, y, "HR / Authorized Signatory")

    c.line(40, y + 18, 140, y + 18)
    c.line(260, y + 18, 360, y + 18)
    c.line(470, y + 18, 570, y + 18)

    y -= 34

    c.setFont("Helvetica-Oblique", 7.5)
    c.drawCentredString(
        width / 2,
        y,
        "This is a computer-generated Full & Final settlement sheet. Signature may not be required if digitally approved."
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

        conn.commit()

        flash(
            f"Full & Final settlement deleted successfully for {row['employee_name']}.",
            "success"
        )

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

    conn = get_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            DELETE FROM payroll_history
            WHERE company_id = ?
              AND month = ?
        """, (company_id, month))

        deleted_count = cur.rowcount

        conn.commit()

        if deleted_count > 0:
            flash(f"Payroll deleted successfully for {month}. Leave records and leave balances were not changed.", "success")
        else:
            flash(f"No payroll records found for {month}.", "warning")

    except Exception as e:
        conn.rollback()
        flash(f"Error while deleting payroll: {str(e)}", "danger")

    finally:
        conn.close()

    return redirect(url_for("payroll_history"))


# ---------------------------
# PAYSLIP PDF
# ---------------------------
def generate_payslip(row):
    os.makedirs(PAYSLIP_FOLDER, exist_ok=True)
    file_name = f"{row['emp_code']}_{row['month']}.pdf"
    file_path = os.path.join(PAYSLIP_FOLDER, file_name)

    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter

    def money(value):
        try:
            return f"Rs. {int(round(float(value or 0)))}"
        except Exception:
            return "Rs. 0"

    def clean_value(value, default="-"):
        if value is None:
            return default

        value = str(value).strip()

        if value == "":
            return default

        if value.lower() in ["nan", "none", "null"]:
            return default

        return value

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

    # Company details
    company_name = clean_value(row_value("company_name"), "SMART HIRE AI PAYROLL")
    company_address = clean_value(row_value("company_address"), "")
    company_email = clean_value(row_value("company_email"), "")
    company_phone = clean_value(row_value("company_phone"), "")

    company_contact_parts = []

    if company_address not in ["", "-"]:
        company_contact_parts.append(company_address)

    if company_email not in ["", "-"]:
        company_contact_parts.append(company_email)

    if company_phone not in ["", "-"]:
        company_contact_parts.append(company_phone)

    company_contact_line = " | ".join(company_contact_parts) if company_contact_parts else "-"

    # Employee details
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

    # Attendance + leave details
    working_days = row_value("attendance_working_days", row_value("working_days", 0))
    present_days = row_value("attendance_present_days", row_value("present_days", 0))
    overtime_hours = row_value("attendance_overtime_hours", row_value("overtime_hours", 0))

    paid_leave_days = row_value("paid_leave_days", 0)
    lwp_days = row_value("lwp_days", 0)
    payable_days = row_value("payable_days", 0)
    lwp_deduction = row_value("lwp_deduction", 0)

    y = height - 35

    c.setFillColorRGB(0.10, 0.17, 0.28)
    c.rect(35, y - 55, width - 70, 55, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)

    c.setFont("Helvetica-Bold", 17)
    c.drawCentredString(width / 2, y - 20, company_name.upper())

    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, y - 38, "Corporate Payroll Management System")
    c.drawCentredString(width / 2, y - 50, company_contact_line)

    y -= 85
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, f"SALARY SLIP - {clean_value(row_value('month'))}")
    y -= 25

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Employee Information")
    y -= 10
    c.rect(40, y - 105, width - 80, 105)

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
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(55, yy, f"{label}:")
        c.setFont("Helvetica", 8.5)
        c.drawString(150, yy, clean_value(value))
        yy -= 18

    yy = y - 22
    for label, value in right:
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(330, yy, f"{label}:")
        c.setFont("Helvetica", 8.5)
        c.drawString(415, yy, clean_value(value))
        yy -= 18

    y -= 130

    # Updated Attendance & Leave Summary
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Attendance & Leave Summary")
    y -= 10
    c.rect(40, y - 62, width - 80, 62)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(55, y - 15, "Working Days")
    c.drawString(155, y - 15, "Present Days")
    c.drawString(255, y - 15, "Paid Leave")
    c.drawString(355, y - 15, "LWP Days")
    c.drawString(455, y - 15, "Payable Days")

    c.setFont("Helvetica", 8)
    c.drawString(55, y - 30, number_text(working_days))
    c.drawString(155, y - 30, number_text(present_days))
    c.drawString(255, y - 30, number_text(paid_leave_days))
    c.drawString(355, y - 30, number_text(lwp_days))
    c.drawString(455, y - 30, number_text(payable_days))

    c.setFont("Helvetica-Bold", 8)
    c.drawString(55, y - 50, "Overtime Hours")
    c.drawString(255, y - 50, "LWP Deduction")
    c.drawString(455, y - 50, "Pay Month")

    c.setFont("Helvetica", 8)
    c.drawString(155, y - 50, number_text(overtime_hours))
    c.drawString(355, y - 50, money(lwp_deduction))
    c.drawString(520, y - 50, clean_value(row_value("month")))

    y -= 92

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Salary Details")
    y -= 15

    x = 40
    table_w = width - 80
    row_h = 21
    table_y = y
    rows_count = 9

    c.rect(x, table_y - (row_h * rows_count), table_w, row_h * rows_count)
    c.setFillColorRGB(0.12, 0.36, 0.85)
    c.rect(x, table_y - row_h, table_w, row_h, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)

    c.line(x + table_w * 0.28, table_y, x + table_w * 0.28, table_y - row_h * rows_count)
    c.line(x + table_w * 0.50, table_y, x + table_w * 0.50, table_y - row_h * rows_count)
    c.line(x + table_w * 0.78, table_y, x + table_w * 0.78, table_y - row_h * rows_count)

    c.setFillColorRGB(0, 0, 0)
    for i in range(rows_count + 1):
        c.line(x, table_y - row_h * i, x + table_w, table_y - row_h * i)

    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x + 8, table_y - 14, "Earnings")
    c.drawString(x + table_w * 0.28 + 8, table_y - 14, "Amount")
    c.drawString(x + table_w * 0.50 + 8, table_y - 14, "Deductions")
    c.drawString(x + table_w * 0.78 + 8, table_y - 14, "Amount")
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

    c.setFont("Helvetica", 8.5)
    start_y = table_y - row_h - 14

    for i in range(7):
        yy = start_y - row_h * i
        c.drawString(x + 8, yy, earnings[i][0])
        c.drawRightString(x + table_w * 0.50 - 8, yy, money(earnings[i][1]))
        c.drawString(x + table_w * 0.50 + 8, yy, deductions[i][0])
        c.drawRightString(x + table_w - 8, yy, money(deductions[i][1]))

    total_y = table_y - row_h * 8 - 14
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x + 8, total_y, "Gross Earnings")
    c.drawRightString(x + table_w * 0.50 - 8, total_y, money(row_value("gross", 0)))
    c.drawString(x + table_w * 0.50 + 8, total_y, "Net Pay")
    c.drawRightString(x + table_w - 8, total_y, money(row_value("net_pay", 0)))

    y = table_y - row_h * rows_count - 35

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Employer Contributions / CTC")
    y -= 10
    c.rect(40, y - 70, width - 80, 70)

    c.setFont("Helvetica", 8.5)
    c.drawString(55, y - 18, f"PF Employer: {money(row_value('pf_employer', 0))}")
    c.drawString(55, y - 36, f"ESIC Employer: {money(row_value('esi_employer', 0))}")
    c.drawString(55, y - 54, f"Gratuity: {money(row_value('gratuity', 0))}")
    c.drawString(330, y - 18, f"Monthly CTC: {money(row_value('monthly_ctc', 0))}")
    c.drawString(330, y - 36, f"Annual CTC: {money(row_value('annual_ctc', 0))}")
    c.drawString(330, y - 54, f"Bonus CTC: {money(row_value('bonus_ctc', 0))}")

    y -= 100
    c.setFillColorRGB(0.86, 0.96, 0.89)
    c.rect(40, y - 35, width - 80, 35, fill=1, stroke=0)
    c.setFillColorRGB(0, 0.35, 0.15)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(55, y - 22, f"NET PAYABLE: {money(row_value('net_pay', 0))}")

    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica", 8.5)
    y -= 65
    c.drawString(40, y, "Prepared By")
    c.drawString(260, y, "Checked By")
    c.drawString(470, y, "HR / Authorized Signatory")
    c.line(40, y + 18, 140, y + 18)
    c.line(260, y + 18, 360, y + 18)
    c.line(470, y + 18, 570, y + 18)

    y -= 35
    c.setFont("Helvetica-Oblique", 7.5)
    c.drawCentredString(
        width / 2,
        y,
        "This is a computer-generated payslip. Signature may not be required if digitally approved."
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
        flash("Please select month to download payslips.")
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

        GROUP BY p.id
        ORDER BY p.emp_code
    """, (company_id, month))

    rows = cur.fetchall()
    conn.close()

    if not rows:
        flash("No payroll data found for selected month.")
        return redirect(url_for("payroll_history"))

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for row in rows:
            pdf_path = generate_payslip(row)

            emp_code = str(row["emp_code"]).strip()
            employee_name = str(row["employee_name"] or "Employee").strip().replace(" ", "_")

            pdf_name = f"{emp_code}_{employee_name}_{month}.pdf"
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


if __name__ == "__main__":
    app.run(debug=True)
