import os
from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
from datetime import date, datetime
import pandas as pd
import smtplib, ssl
from email.message import EmailMessage
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz
# -------------------- Configuration --------------------
app = Flask(__name__)
DB_NAME = "pod.db"

FACULTY_EMAILS = os.environ.get("FACULTY_EMAILS", "230100.cd@rmkec.ac.in,230089.cd@rmkec.ac.in").split(",")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_PASS = os.environ.get("SENDER_PASS")
HOLIDAYS = os.environ.get("HOLIDAYS", "2025-09-10,2025-09-25").split(",")  # YYYY-MM-DD

# -------------------- Database Setup --------------------
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS students (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT NOT NULL,
                        reg_no TEXT UNIQUE NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS acknowledgements (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        student_id INTEGER,
                        ack_date TEXT,
                        status TEXT,
                        reason TEXT,
                        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(student_id) REFERENCES students(id))''')
        conn.commit()

def preload_students():
    students = [
        ("ADHITHYAN", "7001"), ("AJAY", "7002"), ("AKASH","7003"), ("ASHREEN FATHIMA","7004"),
        ("ASUWATHRAM","7005"), ("ASWITHA","7006"), ("BHARATH KUMAR","7007"), ("DHARSHINI","7008"),
        ("FATHIMA NADHEERA","7009"), ("SAI REDDY","7010"), ("VENKATESH","7011"), ("GUNANITHI","7012"),
        ("VEERA SEKHAR","7013"), ("GURU SANTHOSH","7014"), ("INBARASAN K","7015"), ("JAFRIN MERCY","7016"),
        ("PRATHYUMNAN","7017"), ("KEERTHANA","7018"), ("KOWSHIKA","7019"), ("KIRTHISRI","7020"),
        ("LAVANYA","7021"), ("LITHICKA","7022"), ("LOCHAN","7023"), ("KIRTHI DHARSHAN","7024"),
        ("MUKESH","7025"), ("NITHISH","7026"), ("NITHISHWARAN","7027"), ("KOUSHIK REDDY","7028"),
        ("PAVITHRA","7029"), ("PRADEEPTHA","7030"), ("NAVEEN","7061")
        # Add remaining students...
    ]
    with sqlite3.connect(DB_NAME) as conn:
        c = conn.cursor()
        c.execute("DELETE FROM students")
        for s in students:
            try:
                c.execute("INSERT INTO students (name, reg_no) VALUES (?, ?)", s)
            except:
                pass
        conn.commit()

# -------------------- Routes --------------------
@app.route("/", methods=["GET", "POST"])
def index():
    today = str(date.today())

    if request.method == "POST":
        student_id = request.form.get("student_id")
        completed = request.form.get("completed")
        reason = request.form.get("reason")

        status = "Completed" if completed == "on" else "Not Completed"

        with sqlite3.connect(DB_NAME) as conn:
            c = conn.cursor()
            c.execute("SELECT id FROM acknowledgements WHERE student_id=? AND ack_date=?",
                      (student_id, today))
            existing = c.fetchone()
            if existing:
                c.execute("UPDATE acknowledgements SET status=?, reason=? WHERE id=?",
                          (status, reason, existing[0]))
            else:
                c.execute("INSERT INTO acknowledgements (student_id, ack_date, status, reason) VALUES (?, ?, ?, ?)",
                          (student_id, today, status, reason))
            conn.commit()

        return redirect(url_for("index"))

    with sqlite3.connect(DB_NAME) as conn:
        c = conn.cursor()
        c.execute("""SELECT s.id, s.reg_no, s.name,
                     COALESCE(a.status, 'Not Completed'),
                     COALESCE(a.reason, '')
                     FROM students s
                     LEFT JOIN acknowledgements a
                     ON s.id = a.student_id AND a.ack_date=?""", (today,))
        students = c.fetchall()

    return render_template("index.html", students=students, today=today)

@app.route("/report")
def report():
    today = str(date.today())
    filename = generate_report(today)
    return send_file(filename, as_attachment=True)

# -------------------- Report & Email --------------------
def generate_report(today):
    with sqlite3.connect(DB_NAME) as conn:
        query = """SELECT ? AS Date, s.reg_no AS 'Reg No', s.name AS 'Name',
                   COALESCE(a.status, 'Not Completed') AS 'Completion',
                   COALESCE(a.reason, '') AS 'Reason'
                   FROM students s
                   LEFT JOIN acknowledgements a
                   ON s.id = a.student_id AND a.ack_date=?"""
        df = pd.read_sql_query(query, conn, params=(today, today))

    filename = f"PoD_Report_{today}.xlsx"
    df.to_excel(filename, index=False)
    return filename

def send_report_via_email():
    today = str(date.today())
    filename = generate_report(today)

    if not SENDER_EMAIL or not SENDER_PASS:
        print("[ERROR] Email credentials not set!")
        return

    msg = EmailMessage()
    msg["Subject"] = f"PoD Report - {today}"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(FACULTY_EMAILS)
    msg.set_content("Dear Faculty/HOD,\n\nPlease find attached today's PoD completion report.\n\nRegards,\nPoD Portal")

    with open(filename, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(SENDER_EMAIL, SENDER_PASS)
            server.send_message(msg)
        print(f"[INFO] Report sent successfully on {today}")
    except Exception as e:
        print(f"[ERROR] Failed to send email: {e}")

# -------------------- Scheduler --------------------
def send_report_if_workday():
    today_str = datetime.now().strftime("%Y-%m-%d")
    weekday = datetime.now().weekday()  # Monday=0
    if weekday < 5 and today_str not in HOLIDAYS:
        send_report_via_email()
    else:
        print(f"[INFO] Skipped sending report on {today_str} (weekend/holiday)")

def start_scheduler():
    scheduler = BackgroundScheduler(timezone=pytz.timezone("Asia/Kolkata"))  # IST timezone
    scheduler.add_job(send_report_if_workday,CronTrigger(hour=12, minute=50, timezone=pytz.timezone("Asia/Kolkata")))
    scheduler.start()

# -------------------- Main --------------------
if __name__ == "__main__":
    init_db()
    preload_students()
    start_scheduler()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
