import os
import sqlite3
import datetime
from functools import wraps

import pandas as pd
import razorpay

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file, jsonify
)

from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

app = Flask(__name__)
app.secret_key = "change-this-secret-key"

RAZORPAY_KEY_ID = "rzp_test_SfKO3IFwsgnWhC"
RAZORPAY_SECRET = "F7CyUtorZhRe4q8YXtIefnv9"

razorpay_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_SECRET))

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


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT NOT NULL
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

    conn.commit()
    conn.close()


init_db()


# ---------------------------
# Helpers
# ---------------------------
def rupee(value):
    return int(round(float(value or 0)))


def money_str(value):
    return str(int(round(float(value or 0))))


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def current_company_id():
    return session.get("company_id")


def is_pro_user():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT * FROM subscriptions
        WHERE company_id = ? AND status = 'active'
        ORDER BY id DESC
    """, (current_company_id(),))

    sub = cur.fetchone()
    conn.close()

    return sub is not None

def get_active_plan():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT plan_name, status, start_date, end_date
        FROM subscriptions
        WHERE company_id = ? AND status = 'active'
        ORDER BY id DESC
        LIMIT 1
    """, (current_company_id(),))

    plan = cur.fetchone()
    conn.close()
    return plan


def month_only(month_str: str) -> str:
    month_str = str(month_str).strip()
    if "-" in month_str:
        return month_str.split("-")[-1].zfill(2)
    return month_str.zfill(2)


# ---------------------------
# COMPLIANCE RULES
# ---------------------------
def calculate_professional_tax_maharashtra(gross_salary, gender, payroll_month_mm):
    gross_salary = float(gross_salary or 0)
    gender = str(gender or "male").strip().lower()
    payroll_month_mm = month_only(payroll_month_mm)

    if gender == "female":
        if gross_salary <= 25000:
            return 0.0
        return 300.0 if payroll_month_mm == "02" else 200.0

    if gross_salary <= 7500:
        return 0.0
    elif gross_salary <= 10000:
        return 175.0
    else:
        return 300.0 if payroll_month_mm == "02" else 200.0


def calculate_lwf_maharashtra(payroll_month_mm):
    payroll_month_mm = month_only(payroll_month_mm)
    if payroll_month_mm in ["06", "12"]:
        return {"employee": 25.0, "employer": 75.0}
    return {"employee": 0.0, "employer": 0.0}


def calculate_annual_tax_new_regime(annual_taxable_income):
    tax = 0.0
    slabs = [
        (400000, 0.00),
        (800000, 0.05),
        (1200000, 0.10),
        (1600000, 0.15),
        (2000000, 0.20),
        (2400000, 0.25),
        (float("inf"), 0.30),
    ]

    previous_limit = 0
    for limit, rate in slabs:
        if annual_taxable_income > previous_limit:
            taxable_part = min(annual_taxable_income, limit) - previous_limit
            tax += taxable_part * rate
            previous_limit = limit
        else:
            break

    return round(tax, 2)


def calculate_annual_tax_old_regime(annual_taxable_income):
    if annual_taxable_income <= 250000:
        tax = 0
    elif annual_taxable_income <= 500000:
        tax = (annual_taxable_income - 250000) * 0.05
    elif annual_taxable_income <= 1000000:
        tax = 12500 + (annual_taxable_income - 500000) * 0.20
    else:
        tax = 112500 + (annual_taxable_income - 1000000) * 0.30

    return round(tax, 2)


def calculate_monthly_tds(
    monthly_gross,
    employee_pf_monthly=0,
    regime="new",
    other_annual_deductions=0,
    months_remaining=12
):
    monthly_gross = float(monthly_gross or 0)
    employee_pf_monthly = float(employee_pf_monthly or 0)
    other_annual_deductions = float(other_annual_deductions or 0)
    months_remaining = max(int(months_remaining or 12), 1)

    annual_salary = monthly_gross * 12
    annual_pf = employee_pf_monthly * 12
    regime = str(regime or "new").strip().lower()

    if regime == "old":
        standard_deduction = 50000
        annual_taxable_income = annual_salary - standard_deduction - annual_pf - other_annual_deductions
        annual_tax = calculate_annual_tax_old_regime(max(annual_taxable_income, 0))
    else:
        standard_deduction = 75000
        annual_taxable_income = annual_salary - standard_deduction
        annual_tax = calculate_annual_tax_new_regime(max(annual_taxable_income, 0))

    monthly_tds = round(annual_tax / months_remaining, 2)

    return {
        "annual_salary": round(annual_salary, 2),
        "annual_taxable_income": round(max(annual_taxable_income, 0), 2),
        "annual_tax": round(annual_tax, 2),
        "monthly_tds": monthly_tds
    }


def calculate_bonus_logic(basic, payroll_month):
    bonus_ctc = rupee(float(basic or 0) * 0.0833)
    payroll_month = month_only(payroll_month)
    diwali_month = "11"
    if payroll_month == diwali_month:
        festival_bonus = rupee(bonus_ctc * 12)
    else:
        festival_bonus = 0
    return bonus_ctc, festival_bonus


# ---------------------------
# PAYROLL CALCULATION
# ---------------------------
def calculate_payroll_annexure(
    monthly_salary,
    working_days,
    present_days,
    overtime_hours,
    bonus,
    manual_deduction,
    gender,
    payroll_month,
    tax_regime="new",
    other_annual_deductions=0,
    special_allowance_input=0
):
    monthly_salary = float(monthly_salary or 0)
    working_days = int(working_days or 0)
    present_days = int(present_days or 0)
    overtime_hours = float(overtime_hours or 0)
    bonus = rupee(bonus or 0)
    manual_deduction = rupee(manual_deduction or 0)
    special_allowance_input = rupee(special_allowance_input or 0)

    attendance_ratio = (present_days / working_days) if working_days > 0 else 0

    prorated_gross = rupee(monthly_salary * attendance_ratio)

    per_day = monthly_salary / 30 if monthly_salary else 0
    per_hour = per_day / 8 if per_day else 0
    overtime_amount = rupee(per_hour * overtime_hours * 2)

    gross = rupee(prorated_gross + overtime_amount + bonus)

    basic = rupee(gross * 0.40)
    da = rupee(gross * 0.10)
    hra = rupee(gross * 0.20)
    special_allowance = rupee(special_allowance_input)
    other_allowance = rupee(gross - (basic + da + hra + special_allowance))

    pf_wages = basic + da
    pf_employee = rupee(min(pf_wages * 0.12, 1800))
    pf_employer = rupee(min(pf_wages * 0.12, 1800))

    if gross <= 21000:
        esi_employee = rupee(gross * 0.0075)
        esi_employer = rupee(gross * 0.0325)
    else:
        esi_employee = 0
        esi_employer = 0

    professional_tax = rupee(
        calculate_professional_tax_maharashtra(gross, gender, payroll_month)
    )

    lwf = calculate_lwf_maharashtra(payroll_month)
    lwf_employee = rupee(lwf["employee"])
    lwf_employer = rupee(lwf["employer"])

    tds_data = calculate_monthly_tds(
        monthly_gross=gross,
        employee_pf_monthly=pf_employee,
        regime=tax_regime,
        other_annual_deductions=other_annual_deductions,
        months_remaining=12
    )
    tds = rupee(tds_data["monthly_tds"])

    bonus_ctc, festival_bonus = calculate_bonus_logic(basic, payroll_month)

    total_deductions = rupee(
        esi_employee +
        professional_tax +
        pf_employee +
        lwf_employee +
        tds +
        manual_deduction
    )

    gratuity = rupee(basic * 0.0481)

    total_contributions = rupee(
        esi_employer +
        pf_employer +
        gratuity +
        lwf_employer
    )

    net_pay = rupee(gross - total_deductions + festival_bonus)
    monthly_ctc = rupee(gross + total_contributions + bonus_ctc)
    annual_ctc = rupee(monthly_ctc * 12)

    return {
        "basic": basic,
        "da": da,
        "hra": hra,
        "special_allowance": special_allowance,
        "other_allowance": other_allowance,
        "gross": gross,

        "esi_employee": esi_employee,
        "professional_tax": professional_tax,
        "pf_employee": pf_employee,
        "lwf_employee": lwf_employee,
        "tds": tds,
        "manual_deduction": manual_deduction,
        "total_deductions": total_deductions,

        "esi_employer": esi_employer,
        "pf_employer": pf_employer,
        "gratuity": gratuity,
        "bonus_ctc": bonus_ctc,
        "festival_bonus": festival_bonus,
        "lwf_employer": lwf_employer,
        "total_contributions": total_contributions,

        "net_pay": net_pay,
        "monthly_ctc": monthly_ctc,
        "annual_ctc": annual_ctc,

        "overtime_hours": rupee(overtime_hours),
        "overtime_amount": overtime_amount
    }


def generate_payslip(payroll_row):
    filename = f"{payroll_row['emp_code']}_{payroll_row['month']}.pdf"
    file_path = os.path.join(PAYSLIP_FOLDER, filename)

    c = canvas.Canvas(file_path, pagesize=letter)

    c.setFont("Helvetica-Bold", 18)
    c.drawString(170, 760, "SmartHire Payroll Payslip")

    c.setFont("Helvetica", 11)
    c.drawString(60, 735, f"Employee Name: {payroll_row['employee_name']}")
    c.drawString(60, 718, f"Employee Code: {payroll_row['emp_code']}")
    c.drawString(60, 701, f"Role: {payroll_row['role']}")
    c.drawString(60, 684, f"Department: {payroll_row['department']}")
    c.drawString(60, 667, f"Gender: {payroll_row['gender']}")
    c.drawString(60, 650, f"Month: {payroll_row['month']}")

    y = 620

    def line(label, value):
        nonlocal y
        c.drawString(70, y, label)
        c.drawString(330, y, str(value))
        y -= 16

    c.setFont("Helvetica-Bold", 13)
    c.drawString(60, y, "Earnings")
    y -= 22
    c.setFont("Helvetica", 10)

    line("Basic", f"₹{money_str(payroll_row['basic'])}")
    line("DA", f"₹{money_str(payroll_row['da'])}")
    line("HRA", f"₹{money_str(payroll_row['hra'])}")
    line("Special Allowance", f"₹{money_str(payroll_row['special_allowance'])}")
    line("Other Allowance", f"₹{money_str(payroll_row['other_allowance'])}")
    line("Overtime Amount", f"₹{money_str(payroll_row['overtime_amount'])}")
    line("Gross", f"₹{money_str(payroll_row['gross'])}")

    y -= 8
    c.setFont("Helvetica-Bold", 13)
    c.drawString(60, y, "Deductions")
    y -= 22
    c.setFont("Helvetica", 10)

    line("ESI Employee", f"₹{money_str(payroll_row['esi_employee'])}")
    line("Professional Tax", f"₹{money_str(payroll_row['professional_tax'])}")
    line("PF Employee", f"₹{money_str(payroll_row['pf_employee'])}")
    line("LWF Employee", f"₹{money_str(payroll_row['lwf_employee'])}")
    line("TDS", f"₹{money_str(payroll_row['tds'])}")
    line("Manual Deduction", f"₹{money_str(payroll_row['manual_deduction'])}")
    line("Total Deductions", f"₹{money_str(payroll_row['total_deductions'])}")

    y -= 8
    c.setFont("Helvetica-Bold", 13)
    c.drawString(60, y, "Employer Contributions")
    y -= 22
    c.setFont("Helvetica", 10)

    line("ESI Employer", f"₹{money_str(payroll_row['esi_employer'])}")
    line("PF Employer", f"₹{money_str(payroll_row['pf_employer'])}")
    line("Gratuity", f"₹{money_str(payroll_row['gratuity'])}")
    line("Bonus CTC", f"₹{money_str(payroll_row['bonus_ctc'])}")
    line("Festival Bonus Paid", f"₹{money_str(payroll_row['festival_bonus'])}")
    line("LWF Employer", f"₹{money_str(payroll_row['lwf_employer'])}")
    line("Total Contributions", f"₹{money_str(payroll_row['total_contributions'])}")

    y -= 8
    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, y, f"Net Pay: ₹{money_str(payroll_row['net_pay'])}")
    y -= 18
    c.drawString(60, y, f"Monthly CTC: ₹{money_str(payroll_row['monthly_ctc'])}")
    y -= 18
    c.drawString(60, y, f"Annual CTC: ₹{money_str(payroll_row['annual_ctc'])}")

    c.save()
    return file_path


# ---------------------------
# AUTH
# ---------------------------
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        company_name = request.form["company_name"].strip()
        full_name = request.form["full_name"].strip()
        username = request.form["username"].strip()
        password = request.form["password"].strip()

        conn = get_db()
        cur = conn.cursor()

        try:
            cur.execute("INSERT INTO companies (company_name) VALUES (?)", (company_name,))
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


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------
# DASHBOARD
# ---------------------------
@app.route("/")
@login_required
def dashboard():
    company_id = current_company_id()
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT company_name FROM companies WHERE id = ?", (company_id,))
    company = cur.fetchone()

    cur.execute("SELECT COUNT(*) as count FROM employees WHERE company_id = ?", (company_id,))
    employee_count = cur.fetchone()["count"]

    cur.execute("SELECT COUNT(*) as count FROM attendance WHERE company_id = ?", (company_id,))
    attendance_count = cur.fetchone()["count"]

    cur.execute("SELECT COUNT(*) as count FROM payroll_history WHERE company_id = ?", (company_id,))
    payroll_count = cur.fetchone()["count"]

    conn.close()

    active_plan = get_active_plan()

    return render_template(
        "dashboard.html",
        company_name=company["company_name"],
        employee_count=employee_count,
        attendance_count=attendance_count,
        payroll_count=payroll_count,
        active_plan=active_plan
    )

@app.route("/pricing")
@login_required
def pricing():
    active_plan = get_active_plan()
    return render_template(
        "pricing.html",
        razorpay_key=RAZORPAY_KEY_ID,
        active_plan=active_plan
    )

@app.route("/create-order", methods=["POST"])
@login_required
def create_order():
    data = request.get_json()
    amount = int(data.get("amount", 0)) * 100  # convert to paise

    if amount <= 0:
        return jsonify({"status": "failed", "message": "Invalid amount"}), 400

    order = razorpay_client.order.create({
        "amount": amount,
        "currency": "INR",
        "payment_capture": 1
    })

    return jsonify(order)

@app.route("/verify-payment", methods=["POST"])
@app.route("/payments")
@login_required
def payments():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT amount, payment_id, order_id, status, created_at
        FROM payments
        WHERE company_id = ?
        ORDER BY id DESC
    """, (current_company_id(),))

    rows = cur.fetchall()
    conn.close()

    return render_template("payments.html", rows=rows)
@login_required
def verify_payment():
    data = request.get_json()

    try:
        razorpay_client.utility.verify_payment_signature({
            "razorpay_order_id": data["razorpay_order_id"],
            "razorpay_payment_id": data["razorpay_payment_id"],
            "razorpay_signature": data["razorpay_signature"]
        })

        conn = get_db()
        cur = conn.cursor()

        # activate plan
        cur.execute("""
            INSERT INTO subscriptions (company_id, plan_name, status, start_date)
            VALUES (?, ?, ?, ?)
        """, (
            current_company_id(),
            "pro",
            "active",
            str(datetime.date.today())
        ))

        # store payment
        cur.execute("""
            INSERT INTO payments (company_id, amount, payment_id, order_id, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            current_company_id(),
            999,
            data["razorpay_payment_id"],
            data["razorpay_order_id"],
            "success",
            str(datetime.datetime.now())
        ))

        conn.commit()
        conn.close()

        return jsonify({"status": "success"})

    except Exception as e:
        return jsonify({"status": "failed", "message": str(e)}), 400


# ---------------------------
# EMPLOYEE MASTER UPLOAD
# ---------------------------
@app.route("/upload-employees", methods=["GET", "POST"])
@login_required
def upload_employees():
    if request.method == "POST":
        file = request.files["file"]
        if not file or file.filename == "":
            flash("Please select a file.")
            return redirect(url_for("upload_employees"))

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        try:
            if file.filename.endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path, engine="openpyxl")

            conn = get_db()
            cur = conn.cursor()

            for _, row in df.iterrows():
                try:
                    cur.execute("""
                        INSERT OR REPLACE INTO employees
                        (company_id, emp_code, employee_name, role, department, gender, monthly_salary, tax_regime, other_annual_deductions, special_allowance)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        current_company_id(),
                        str(row["emp_code"]).strip(),
                        str(row["employee_name"]).strip(),
                        str(row["role"]).strip(),
                        str(row.get("department", "")).strip(),
                        str(row.get("gender", "male")).strip().lower(),
                        float(row["monthly_salary"]),
                        str(row.get("tax_regime", "new")).strip().lower(),
                        float(row.get("other_annual_deductions", 0) or 0),
                        float(row.get("special_allowance", 0) or 0)
                    ))
                except Exception:
                    continue

            conn.commit()
            conn.close()
            flash("Employee master uploaded successfully.")
            return redirect(url_for("employees_list"))

        except Exception as e:
            flash(f"Upload failed: {e}")

    return render_template("upload_employees.html")


@app.route("/employees")
@login_required
def employees_list():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM employees WHERE company_id = ? ORDER BY id DESC", (current_company_id(),))
    employees = cur.fetchall()
    conn.close()
    return render_template("employees.html", employees=employees)


# ---------------------------
# ATTENDANCE UPLOAD
# ---------------------------
@app.route("/upload-attendance", methods=["GET", "POST"])
@login_required
def upload_attendance():
    if request.method == "POST":
        file = request.files["file"]
        if not file or file.filename == "":
            flash("Please select a file.")
            return redirect(url_for("upload_attendance"))

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        try:
            if file.filename.endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path, engine="openpyxl")

            conn = get_db()
            cur = conn.cursor()

            for _, row in df.iterrows():
                try:
                    cur.execute("""
                        INSERT INTO attendance
                        (company_id, emp_code, month, working_days, present_days, overtime_hours, bonus, manual_deduction)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        current_company_id(),
                        str(row["emp_code"]).strip(),
                        str(row["month"]).strip(),
                        int(row["working_days"]),
                        int(row["present_days"]),
                        float(row.get("overtime_hours", 0) or 0),
                        float(row.get("bonus", 0) or 0),
                        float(row.get("manual_deduction", 0) or 0)
                    ))
                except Exception:
                    continue

            conn.commit()
            conn.close()
            flash("Attendance uploaded successfully.")
            return redirect(url_for("attendance_list"))

        except Exception as e:
            flash(f"Upload failed: {e}")

    return render_template("upload_attendance.html")


@app.route("/attendance")
@login_required
def attendance_list():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM attendance
        WHERE company_id = ?
        ORDER BY id DESC
    """, (current_company_id(),))
    rows = cur.fetchall()
    conn.close()
    return render_template("attendance.html", rows=rows)


# ---------------------------
# RUN PAYROLL
# ---------------------------
@app.route("/run-payroll", methods=["GET", "POST"])
@login_required
def run_payroll():
    if request.method == "POST":
        payroll_month = request.form["month"].strip()

        conn = get_db()
        cur = conn.cursor()

        cur.execute("""
            SELECT
                a.*,
                e.employee_name,
                e.role,
                e.department,
                e.gender,
                e.monthly_salary,
                e.tax_regime,
                e.other_annual_deductions,
                e.special_allowance
            FROM attendance a
            JOIN employees e
                ON a.emp_code = e.emp_code
               AND a.company_id = e.company_id
            WHERE a.company_id = ? AND a.month = ?
        """, (current_company_id(), payroll_month))

        rows = cur.fetchall()

        if not rows:
            conn.close()
            flash("No attendance data found for this month.")
            return redirect(url_for("run_payroll"))

        for row in rows:
            calc = calculate_payroll_annexure(
                monthly_salary=row["monthly_salary"],
                working_days=row["working_days"],
                present_days=row["present_days"],
                overtime_hours=row["overtime_hours"],
                bonus=row["bonus"],
                manual_deduction=row["manual_deduction"],
                gender=row["gender"],
                payroll_month=row["month"],
                tax_regime=row["tax_regime"],
                other_annual_deductions=row["other_annual_deductions"],
                special_allowance_input=row["special_allowance"]
            )

            cur.execute("""
                INSERT INTO payroll_history (
                    company_id, emp_code, employee_name, role, department, gender, month,
                    monthly_salary, basic, da, hra, special_allowance, other_allowance, gross,
                    esi_employee, professional_tax, pf_employee, lwf_employee, tds,
                    manual_deduction, total_deductions,
                    esi_employer, pf_employer, gratuity, bonus_ctc, festival_bonus, lwf_employer,
                    total_contributions, net_pay, monthly_ctc, annual_ctc,
                    overtime_hours, overtime_amount, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                current_company_id(),
                row["emp_code"],
                row["employee_name"],
                row["role"],
                row["department"],
                row["gender"],
                payroll_month,
                row["monthly_salary"],
                calc["basic"],
                calc["da"],
                calc["hra"],
                calc["special_allowance"],
                calc["other_allowance"],
                calc["gross"],

                calc["esi_employee"],
                calc["professional_tax"],
                calc["pf_employee"],
                calc["lwf_employee"],
                calc["tds"],
                calc["manual_deduction"],
                calc["total_deductions"],

                calc["esi_employer"],
                calc["pf_employer"],
                calc["gratuity"],
                calc["bonus_ctc"],
                calc["festival_bonus"],
                calc["lwf_employer"],
                calc["total_contributions"],

                calc["net_pay"],
                calc["monthly_ctc"],
                calc["annual_ctc"],

                calc["overtime_hours"],
                calc["overtime_amount"],
                str(datetime.datetime.now())
            ))

        conn.commit()
        conn.close()

        flash("Payroll run completed successfully.")
        return redirect(url_for("payroll_history"))

    return render_template("run_payroll.html")


# ---------------------------
# PAYROLL HISTORY + FILTERS
# ---------------------------
@app.route("/payroll-history")
@login_required
def payroll_history():
    month = request.args.get("month", "").strip()
    department = request.args.get("department", "").strip()

    conn = get_db()
    cur = conn.cursor()

    query = """
        SELECT * FROM payroll_history
        WHERE company_id = ?
    """
    params = [current_company_id()]

    if month:
        query += " AND month = ?"
        params.append(month)

    if department:
        query += " AND department = ?"
        params.append(department)

    query += " ORDER BY id DESC"

    cur.execute(query, tuple(params))
    rows = cur.fetchall()

    cur.execute("""
        SELECT DISTINCT department FROM payroll_history
        WHERE company_id = ? AND department IS NOT NULL AND department != ''
        ORDER BY department
    """, (current_company_id(),))
    departments = cur.fetchall()

    conn.close()

    return render_template(
        "payroll_history.html",
        rows=rows,
        departments=departments,
        selected_month=month,
        selected_department=department
    )


# ---------------------------
# EXPORT EXCEL
# ---------------------------
@app.route("/export-excel")
@login_required
def export_excel():
    if not is_pro_user():
        flash("Upgrade to PRO to use Excel export")
        return redirect(url_for("pricing"))
    month = request.args.get("month", "").strip()
    department = request.args.get("department", "").strip()

    conn = get_db()
    query = """
        SELECT
            emp_code, employee_name, role, department, gender, month, monthly_salary,
            basic, da, hra, special_allowance, other_allowance, gross,
            esi_employee, professional_tax, pf_employee, lwf_employee, tds,
            manual_deduction, total_deductions,
            esi_employer, pf_employer, gratuity, bonus_ctc, festival_bonus, lwf_employer,
            total_contributions, net_pay, monthly_ctc, annual_ctc,
            overtime_hours, overtime_amount, created_at
        FROM payroll_history
        WHERE company_id = ?
    """
    params = [current_company_id()]

    if month:
        query += " AND month = ?"
        params.append(month)

    if department:
        query += " AND department = ?"
        params.append(department)

    query += " ORDER BY id DESC"

    df = pd.read_sql_query(query, conn, params=tuple(params))
    conn.close()

    money_columns = [
        "monthly_salary", "basic", "da", "hra", "special_allowance",
        "other_allowance", "gross", "esi_employee", "professional_tax",
        "pf_employee", "lwf_employee", "tds", "manual_deduction",
        "total_deductions", "esi_employer", "pf_employer", "gratuity",
        "bonus_ctc", "festival_bonus", "lwf_employer",
        "total_contributions", "net_pay", "monthly_ctc", "annual_ctc",
        "overtime_hours", "overtime_amount"
    ]

    for col in money_columns:
        if col in df.columns:
            df[col] = df[col].fillna(0).round().astype(int)

    file_name = "payroll_export.xlsx"
    df.to_excel(file_name, index=False)

    return send_file(file_name, as_attachment=True)


# ---------------------------
# PAYSLIP PDF
# ---------------------------
@app.route("/download-payslip/<int:payroll_id>")
@login_required
def download_payslip(payroll_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM payroll_history
        WHERE id = ? AND company_id = ?
    """, (payroll_id, current_company_id()))
    row = cur.fetchone()
    conn.close()

    if not row:
        flash("Payslip not found.")
        return redirect(url_for("payroll_history"))

    pdf_path = generate_payslip(row)
    return send_file(pdf_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)