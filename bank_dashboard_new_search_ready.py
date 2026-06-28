"""
Bank Submit History — Streamlit Dashboard (Smart Upgrade v2)
=============================================================
Features:
  - Login system (admin / manager / sales roles)
  - Smart search (natural language sidebar)
  - Alerts panel (overdue, due-soon, anomaly)
  - Forecasting (linear trend, next-month prediction)
  - Due Date Tracker tab
  - Sales Person Leaderboard (gamified)
  - Enhanced PDF reports
  - All original tabs preserved
"""

import os, glob, shutil, tempfile, re, math
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from io import BytesIO

try:
    import tomllib
except ImportError:
    try:
        import tomli as tomllib
    except ImportError:
        tomllib = None

try:
    from reportlab.lib.pagesizes import A3, landscape as _landscape
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, PageBreak, Spacer
    from reportlab.pdfbase import pdfmetrics
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Bank Submit Dashboard", page_icon="🏦",
                   layout="wide", initial_sidebar_state="expanded")

# ─────────────────────────────────────────────────────────────────────────────
# LOGIN SYSTEM + USER MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────
import json

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(DATA_DIR, "users.json")
ACTIVITY_FILE = os.path.join(DATA_DIR, "activity_log.json")
SESSIONS_FILE = os.path.join(DATA_DIR, "sessions.json")
EMAIL_SCHEDULE_FILE = os.path.join(DATA_DIR, "email_schedule.json")

TAB_MAP = {
    "daily": "Daily Analysis", "overview": "Overview", "weekly": "Weekly Analysis",
    "firm": "Firm & Sales Person", "banks": "Banks", "parties": "Top Parties",
    "payment": "Payment Status", "accept": "Bank Accept Analysis",
    "asm": "Asm Analysis", "due": "Due Date Tracker", "leaderboard": "Leaderboard",
}

DEFAULT_USERS = {
    "admin": {"password": "bank@2026", "role": "admin", "name": "Administrator",
              "email": "admin@company.com", "status": "active",
              "created_at": "2026-01-01", "last_login": "", "login_count": 0,
              "permissions": ["all"], "sales_person": ""},
    "manager": {"password": "manager123", "role": "manager", "name": "Manager",
                "email": "manager@company.com", "status": "active",
                "created_at": "2026-01-01", "last_login": "", "login_count": 0,
                "permissions": ["all"], "sales_person": ""},
    "sales": {"password": "sales123", "role": "sales", "name": "Sales Team",
              "email": "sales@company.com", "status": "active",
              "created_at": "2026-01-01", "last_login": "", "login_count": 0,
              "permissions": list(TAB_MAP.keys()), "sales_person": ""},
}

def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    save_users(DEFAULT_USERS)
    return DEFAULT_USERS.copy()

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)

def load_activity():
    if os.path.exists(ACTIVITY_FILE):
        try:
            with open(ACTIVITY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"logs": []}

def save_activity(data):
    with open(ACTIVITY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def load_sessions():
    if os.path.exists(SESSIONS_FILE):
        try:
            with open(SESSIONS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"sessions": {}}

def save_sessions(data):
    with open(SESSIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

DEFAULT_EMAIL_SCHEDULE = {
    "smtp_host": "smtp.gmail.com", "smtp_port": 587,
    "smtp_user": "", "smtp_pass": "", "use_tls": True,
    "from_name": "Bank Dashboard Reports", "from_email": "",
    "recipients": [], "schedule": "off", "send_day": "Monday",
    "send_time": "09:00", "last_sent": "", "report_title": "Bank Submit Bill Report",
    "auto_enabled": False,
}

def load_email_schedule():
    if os.path.exists(EMAIL_SCHEDULE_FILE):
        try:
            with open(EMAIL_SCHEDULE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_EMAIL_SCHEDULE.items():
                if k not in data:
                    data[k] = v
            return data
        except Exception:
            pass
    return dict(DEFAULT_EMAIL_SCHEDULE)

def save_email_schedule(data):
    with open(EMAIL_SCHEDULE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def send_report_email(pdf_bytes, title="Bank Submit Bill Report"):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders
    cfg = load_email_schedule()
    if not cfg.get("smtp_user") or not cfg.get("recipients"):
        return False, "SMTP not configured or no recipients"
    msg = MIMEMultipart()
    msg["From"] = f"{cfg.get('from_name', 'Dashboard')} <{cfg['from_email']}>"
    msg["To"] = ", ".join(cfg["recipients"])
    msg["Subject"] = f"{title} — {datetime.now().strftime('%d %b %Y %H:%M')}"
    body = f"""Hello,

Please find attached the latest report: {title}
Generated on: {datetime.now().strftime('%d %B %Y at %H:%M')}

This is an automated report from Bank Submit Dashboard.
"""
    msg.attach(MIMEText(body, "plain"))
    part = MIMEBase("application", "pdf")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    filename = f"{title.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    part.add_header("Content-Disposition", f"attachment; filename={filename}")
    msg.attach(part)
    try:
        server = smtplib.SMTP(cfg["smtp_host"], int(cfg["smtp_port"]))
        if cfg.get("use_tls", True):
            server.starttls()
        server.login(cfg["smtp_user"], cfg["smtp_pass"])
        server.sendmail(cfg["from_email"], cfg["recipients"], msg.as_string())
        server.quit()
        cfg["last_sent"] = datetime.now().isoformat()
        save_email_schedule(cfg)
        return True, f"Email sent to {len(cfg['recipients'])} recipient(s)"
    except Exception as e:
        return False, f"Email error: {str(e)}"

def check_auto_email_schedule(pdf_bytes_func, title="Bank Submit Bill Report"):
    cfg = load_email_schedule()
    if not cfg.get("auto_enabled") or cfg.get("schedule", "off") == "off":
        return
    last = cfg.get("last_sent", "")
    now = datetime.now()
    should_send = False
    if cfg["schedule"] == "daily":
        if not last or (now - datetime.fromisoformat(last)).total_seconds() > 86400:
            should_send = True
    elif cfg["schedule"] == "weekly":
        target_day = {"Monday":0,"Tuesday":1,"Wednesday":2,"Thursday":3,"Friday":4,"Saturday":5,"Sunday":6}.get(cfg.get("send_day","Monday"),0)
        if not last:
            should_send = now.weekday() == target_day
        else:
            last_dt = datetime.fromisoformat(last)
            days_since = (now - last_dt).total_seconds() / 86400
            if days_since >= 7 and now.weekday() == target_day:
                should_send = True
    elif cfg["schedule"] == "monthly":
        if not last:
            should_send = now.day == 1
        else:
            last_dt = datetime.fromisoformat(last)
            if now.month != last_dt.month and now.day == 1:
                should_send = True
    if should_send:
        pdf_bytes = pdf_bytes_func()
        if pdf_bytes:
            ok, msg = send_report_email(pdf_bytes, title)
            if ok:
                log_activity("system", "auto_email_sent", msg)
            else:
                log_activity("system", "auto_email_failed", msg)

def register_session(user):
    sid = st.session_state.get("_sid")
    if not sid:
        import hashlib
        sid = hashlib.md5(f"{user}_{datetime.now().timestamp()}".encode()).hexdigest()[:12]
        st.session_state._sid = sid
    sessions = load_sessions()
    sessions["sessions"][sid] = {
        "user": user,
        "login_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "last_active": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_sessions(sessions)
    return sid

def heartbeat_session():
    sid = st.session_state.get("_sid")
    if sid:
        sessions = load_sessions()
        if sid in sessions.get("sessions", {}):
            sessions["sessions"][sid]["last_active"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            save_sessions(sessions)

def kill_session(sid):
    sessions = load_sessions()
    if sid in sessions.get("sessions", {}):
        user = sessions["sessions"][sid].get("user", "unknown")
        del sessions["sessions"][sid]
        save_sessions(sessions)
        log_activity(st.session_state.get("username", "admin"), "force_logout", f"Killed session for: {user}")

def kill_all_sessions(exclude_user=""):
    sessions = load_sessions()
    keep = {k: v for k, v in sessions.get("sessions", {}).items() if v.get("user") == exclude_user}
    sessions["sessions"] = keep
    save_sessions(sessions)

def log_activity(user, action, details=""):
    act = load_activity()
    act["logs"].insert(0, {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user": user, "action": action, "details": details,
    })
    if len(act["logs"]) > 500:
        act["logs"] = act["logs"][:500]
    save_activity(act)

def show_login():
    if st.session_state.get("authenticated"):
        return True
    qp = st.query_params
    if qp.get("user") and qp.get("authed") == "1":
        users = load_users()
        uname = qp["user"].strip().lower()
        if uname in users and users[uname].get("status") == "active":
            st.session_state.authenticated = True
            st.session_state.username = uname
            st.session_state.user_role = users[uname]["role"]
            st.session_state.user_name = users[uname]["name"]
            users[uname]["last_login"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            users[uname]["login_count"] = users[uname].get("login_count", 0) + 1
            save_users(users)
            log_activity(uname, "login", "Auto-login via URL")
            register_session(uname)
            return True
    st.markdown("""
    <style>
    .stApp { background: linear-gradient(135deg, #0b1220 0%, #0d2137 50%, #0b1220 100%); }
    .login-box {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(0,201,167,0.25);
        border-radius: 20px;
        padding: 48px 40px;
        backdrop-filter: blur(12px);
        box-shadow: 0 25px 60px rgba(0,0,0,0.5);
    }
    </style>
    """, unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div class='login-box'>
            <div style='text-align:center; margin-bottom:28px;'>
                <div style='font-size:52px;'>🏦</div>
                <h2 style='color:#00c9a7; margin:8px 0 4px; font-size:26px; letter-spacing:1px;'>Bank Submit Dashboard</h2>
                <p style='color:#556677; font-size:13px; margin:0;'>Please sign in to continue</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        username = st.text_input("Username", placeholder="Enter username", key="login_user")
        password = st.text_input("Password", type="password", placeholder="Enter password", key="login_pass")
        if st.button("Sign In", type="primary", use_container_width=True):
            users = load_users()
            uname = username.strip().lower()
            if uname in users and users[uname]["password"] == password:
                if users[uname].get("status") != "active":
                    st.error("Account is disabled. Contact admin.")
                    return False
                st.session_state.authenticated = True
                st.session_state.username = uname
                st.session_state.user_role = users[uname]["role"]
                st.session_state.user_name = users[uname]["name"]
                st.query_params["user"] = uname
                st.query_params["authed"] = "1"
                users[uname]["last_login"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                users[uname]["login_count"] = users[uname].get("login_count", 0) + 1
                save_users(users)
                log_activity(uname, "login", "Successful login")
                register_session(uname)
                st.rerun()
            else:
                log_activity(username or "unknown", "login_failed", f"Attempt: {username}")
                st.error("Invalid username or password")
        st.markdown("""
        <div style='text-align:center; margin-top:20px; color:#445566; font-size:12px;'>
        Default: <code>admin</code> / <code>bank@2026</code>
        </div>
        """, unsafe_allow_html=True)
    return False

# ─────────────────────────────────────────────────────────────────────────────
# STANDALONE SMART PDF REPORT (used by email button)
# ─────────────────────────────────────────────────────────────────────────────
def smart_overdue_pdf(overdue_df, full_df, period_label="All", generated_at=None):
    if not REPORTLAB_AVAILABLE:
        return None
    from datetime import datetime as _dt
    if generated_at is None:
        generated_at = _dt.now().strftime("%d %b %Y %H:%M")

    def _usd(v):
        try: v = float(v)
        except: return "$0.00"
        if v >= 1e6: return "$" + f"{v/1e6:.2f}" + "M"
        if v >= 1e3: return "$" + f"{v/1e3:.1f}" + "K"
        return "$" + f"{v:.2f}"

    pw, ph = _landscape(A3)
    lm = rm = 40; tm = 50; bm = 45
    uw = pw - lm - rm
    hf = "Helvetica-Bold"; cf = "Helvetica"
    title_clr = colors.HexColor("#0d3f47")
    accent    = colors.HexColor("#00c9a7")
    dark_bg   = colors.HexColor("#0b1220")
    mid_bg    = colors.HexColor("#1a2235")
    header_bg = colors.HexColor("#0d3f47")
    header_fg = colors.white
    text_dark = colors.HexColor("#222222")
    text_mut  = colors.HexColor("#666666")
    overdue_c = colors.HexColor("#ff3b30")
    warn_c    = colors.HexColor("#ff9500")
    orange_c  = colors.HexColor("#ff6b35")
    info_c    = colors.HexColor("#1a8fff")
    safe_c    = colors.HexColor("#00c9a7")
    row_even  = colors.HexColor("#f5f7fa")
    row_odd   = colors.white
    border_c  = colors.HexColor("#d0d5dd")

    def mw(txt, font, sz):
        return pdfmetrics.stringWidth(str(txt), font, sz)

    if len(overdue_df) == 0:
        return None

    firm_overdue = overdue_df.groupby("Firm Name").agg(
        count=("Invoice Value", "size"),
        total_value=("Invoice Value", "sum"),
        max_overdue_days=("Days Until Maturity", "max"),
    ).sort_values("total_value", ascending=False).head(15)

    bank_overdue = overdue_df.groupby("Our Bank").agg(
        count=("Invoice Value", "size"),
        total_value=("Invoice Value", "sum"),
        max_overdue_days=("Days Until Maturity", "max"),
        avg_overdue_days=("Days Until Maturity", "mean"),
    ).sort_values("total_value", ascending=False).head(15)

    _all = full_df if len(full_df) > 0 else overdue_df
    _f_overdue = overdue_df
    _f_due7  = _all[(_all["Days Until Maturity"] <= 0) & (_all["Days Until Maturity"] >= -7)]
    _f_due15 = _all[(_all["Days Until Maturity"] < -7)  & (_all["Days Until Maturity"] >= -15)]
    _f_due30 = _all[(_all["Days Until Maturity"] < -15) & (_all["Days Until Maturity"] >= -30)]
    _f_due60 = _all[(_all["Days Until Maturity"] < -30) & (_all["Days Until Maturity"] >= -60)]

    total_val  = _all["Invoice Value"].sum() if "Invoice Value" in _all.columns else 0
    overdue_val = overdue_df["Invoice Value"].sum() if "Invoice Value" in overdue_df.columns else 0
    due7_val   = _f_due7["Invoice Value"].sum() if "Invoice Value" in _f_due7.columns and len(_f_due7) else 0
    due15_val  = _f_due15["Invoice Value"].sum() if "Invoice Value" in _f_due15.columns and len(_f_due15) else 0
    due30_val  = _f_due30["Invoice Value"].sum() if "Invoice Value" in _f_due30.columns and len(_f_due30) else 0
    due60_val  = _f_due60["Invoice Value"].sum() if "Invoice Value" in _f_due60.columns and len(_f_due60) else 0

    summary_data = {
        "total": len(_all), "total_val": _usd(total_val),
        "overdue": len(_f_overdue), "overdue_val": _usd(overdue_val),
        "due7": len(_f_due7), "due7_val": _usd(due7_val),
        "due15": len(_f_due15), "due15_val": _usd(due15_val),
        "due30": len(_f_due30), "due30_val": _usd(due30_val),
        "due60": len(_f_due60), "due60_val": _usd(due60_val),
    }

    def draw_cover(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(dark_bg); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
        canvas.setFillColor(accent); canvas.roundRect(pw/2 - 200, ph - 180, 400, 8, 4, fill=1, stroke=0)
        canvas.setFillColor(header_fg); canvas.setFont(hf, 32)
        canvas.drawCentredString(pw/2, ph - 250, "DUE DATE TRACKER")
        canvas.setFillColor(accent); canvas.setFont(hf, 18)
        canvas.drawCentredString(pw/2, ph - 285, "Bank Submit History Report")
        canvas.setFillColor(colors.HexColor("#8899aa")); canvas.setFont(cf, 12)
        canvas.drawCentredString(pw/2, ph - 320, f"Report Period: {period_label}")
        canvas.drawCentredString(pw/2, ph - 340, f"Generated: {generated_at}")
        canvas.setFillColor(accent); canvas.roundRect(pw/2 - 200, ph - 370, 400, 2, 1, fill=1, stroke=0)

        box_y = ph - 530; box_h = 110; box_w = (uw - 40) / 4
        kpi_boxes = [
            ("TOTAL RECORDS", str(summary_data['total']), summary_data['total_val'], info_c),
            ("TOTAL VALUE", summary_data['total_val'], str(summary_data['total']) + " Bills", accent),
            ("OVERDUE", str(summary_data['overdue']) + " Bills", summary_data['overdue_val'], overdue_c),
            ("DUE SOON (7d)", str(summary_data['due7']) + " Bills", summary_data['due7_val'], warn_c),
        ]
        for i, (lbl, val, sub, clr) in enumerate(kpi_boxes):
            bx = lm + 20 + i * (box_w + 10)
            canvas.setFillColor(mid_bg); canvas.roundRect(bx, box_y, box_w, box_h, 8, fill=1, stroke=0)
            canvas.setStrokeColor(clr); canvas.setLineWidth(3)
            canvas.line(bx, box_y + box_h, bx + box_w, box_y + box_h)
            canvas.setFillColor(clr); canvas.setFont(hf, 10); canvas.drawString(bx + 15, box_y + box_h - 25, lbl)
            canvas.setFillColor(header_fg); canvas.setFont(hf, 26); canvas.drawString(bx + 15, box_y + 40, val)
            canvas.setFillColor(clr); canvas.setFont(hf, 14); canvas.drawString(bx + 15, box_y + 12, sub)

        canvas.setFillColor(colors.HexColor("#556677")); canvas.setFont(cf, 9)
        canvas.drawCentredString(pw/2, 100, "Prepared by: Smart Dashboard v2.0")
        canvas.drawCentredString(pw/2, 82, "Confidential - For Internal Use Only")
        canvas.setStrokeColor(accent); canvas.setLineWidth(0.5); canvas.line(lm, 60, pw - rm, 60)
        canvas.restoreState()

    def draw_summary(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
        canvas.setFillColor(title_clr); canvas.setFont(hf, 20)
        canvas.drawString(lm, ph - tm - 10, "EXECUTIVE SUMMARY")
        canvas.setStrokeColor(accent); canvas.setLineWidth(2)
        canvas.line(lm, ph - tm - 18, lm + 220, ph - tm - 18)
        canvas.setFillColor(text_mut); canvas.setFont(cf, 10)
        canvas.drawString(lm, ph - tm - 38, f"Report Period: {period_label}  |  Generated: {generated_at}")

        sec_y = ph - tm - 70
        canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
        canvas.drawString(lm, sec_y, "1. Maturity Status Overview")
        canvas.setStrokeColor(border_c); canvas.setLineWidth(0.5)
        canvas.line(lm, sec_y - 5, pw - rm, sec_y - 5)

        card_y = sec_y - 95; card_h = 75; card_w = (uw - 30) / 5
        cards = [
            ("Overdue", summary_data['overdue'], summary_data['overdue_val'], overdue_c),
            ("Due in 7 Days", summary_data['due7'], summary_data['due7_val'], warn_c),
            ("Due in 15 Days", summary_data['due15'], summary_data['due15_val'], orange_c),
            ("Due in 30 Days", summary_data['due30'], summary_data['due30_val'], info_c),
            ("Due in 60 Days", summary_data['due60'], summary_data['due60_val'], safe_c),
        ]
        for i, (lbl, cnt, val, clr) in enumerate(cards):
            cx = lm + i * (card_w + 6)
            canvas.setFillColor(colors.HexColor("#f8f9fa")); canvas.roundRect(cx, card_y, card_w, card_h, 6, fill=1, stroke=0)
            canvas.setFillColor(clr); canvas.roundRect(cx, card_y + card_h - 5, card_w, 5, 2, fill=1, stroke=0)
            canvas.setFont(hf, 9); canvas.drawString(cx + 10, card_y + card_h - 22, lbl)
            canvas.setFillColor(text_dark); canvas.setFont(hf, 22); canvas.drawString(cx + 10, card_y + 30, str(cnt) + " Bills")
            canvas.setFillColor(clr); canvas.setFont(hf, 11); canvas.drawString(cx + 10, card_y + 8, val)

        sec2_y = card_y - 40
        canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
        canvas.drawString(lm, sec2_y, "2. Firm-wise Overdue Summary (Top 15)")
        canvas.setStrokeColor(border_c); canvas.line(lm, sec2_y - 5, pw - rm, sec2_y - 5)

        if len(firm_overdue) > 0:
            firm_header = ["#", "Firm Name", "Bill Count", "Total Bill Value", "Max Overdue", "Risk Level"]
            firm_rows = [firm_header]
            for rank, (firm, frow) in enumerate(firm_overdue.iterrows(), 1):
                max_d = int(frow["max_overdue_days"])
                cnt = int(frow["count"])
                val = _usd(frow["total_value"])
                risk = "CRITICAL" if max_d > 60 else "HIGH" if max_d > 30 else "MEDIUM" if max_d > 15 else "LOW"
                firm_rows.append([str(rank), str(firm)[:30], str(cnt), val, str(max_d) + "d", risk])
            fcol_w = [25, 180, 60, 100, 65, 70]
            ftbl = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
                ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (1, 0), (1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
            ])
            for ri in range(1, len(firm_rows)):
                rv2 = firm_rows[ri][-1]
                if rv2 == "CRITICAL": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); ftbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                elif rv2 == "HIGH": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); ftbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                elif rv2 == "MEDIUM": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), orange_c)
            ft = LongTable(firm_rows, colWidths=fcol_w, hAlign="LEFT")
            ft.setStyle(ftbl)
            tw, th = ft.wrap(520, 200)
            ft.drawOn(canvas, lm, sec2_y - 25 - th)
        else:
            th = 0

        bank_x = lm + 560
        canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
        canvas.drawString(bank_x, sec2_y, "3. Our Bank-wise Overdue Summary")
        canvas.setStrokeColor(border_c); canvas.line(bank_x, sec2_y - 5, bank_x + 520, sec2_y - 5)

        if len(bank_overdue) > 0:
            bheader = ["#", "Our Bank", "Bill Count", "Total Value", "Max Days", "Risk Level"]
            brows = [bheader]
            for rank, (bname, brow) in enumerate(bank_overdue.iterrows(), 1):
                max_d = int(brow["max_overdue_days"])
                cnt = int(brow["count"])
                val = _usd(brow["total_value"])
                risk = "CRITICAL" if max_d > 60 else "HIGH" if max_d > 30 else "MEDIUM" if max_d > 15 else "LOW"
                brows.append([str(rank), str(bname)[:25], str(cnt), val, str(max_d) + "d", risk])
            bcol_w = [25, 120, 55, 85, 55, 70]
            btbl = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
                ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (1, 0), (1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
            ])
            for ri in range(1, len(brows)):
                rv2 = brows[ri][-1]
                if rv2 == "CRITICAL": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); btbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                elif rv2 == "HIGH": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); btbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                elif rv2 == "MEDIUM": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), orange_c)
            bt = LongTable(brows, colWidths=bcol_w, hAlign="LEFT")
            bt.setStyle(btbl)
            btw, bth = bt.wrap(420, 200)
            bt.drawOn(canvas, bank_x, sec2_y - 25 - bth)
        else:
            bth = 0

        sec3_y = sec2_y - 25 - max(th, bth if 'bth' in dir() else 0) - 35
        canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
        canvas.drawString(lm, sec3_y, "4. Key Observations")
        canvas.setStrokeColor(border_c); canvas.line(lm, sec3_y - 5, pw - rm, sec3_y - 5)

        total_overdue_val = overdue_df['Invoice Value'].sum()
        total_all_val = _all['Invoice Value'].sum()
        pct_overdue = (total_overdue_val / total_all_val * 100) if total_all_val > 0 else 0
        top_firm = firm_overdue.index[0] if len(firm_overdue) > 0 else "N/A"
        top_firm_val = _usd(firm_overdue.iloc[0]["total_value"]) if len(firm_overdue) > 0 else "$0"
        top_bank = bank_overdue.index[0] if len(bank_overdue) > 0 else "N/A"
        top_bank_val = _usd(bank_overdue.iloc[0]["total_value"]) if len(bank_overdue) > 0 else "$0"
        critical_firms = len(firm_overdue[firm_overdue["max_overdue_days"] > 60]) if len(firm_overdue) > 0 else 0

        obs = [
            f"OVERALL RISK: {summary_data['overdue']} bills ({pct_overdue:.1f}% of total value) are overdue, totaling {summary_data['overdue_val']}.",
            f"TOP FIRM: {top_firm} leads with {top_firm_val} overdue. Immediate collection follow-up recommended.",
            f"TOP BANK: {top_bank} exposure is {top_bank_val}. Coordinate with relationship manager.",
            f"CRITICAL: {critical_firms} firms overdue >60 days. Escalate to senior management for recovery.",
        ]
        oy = sec3_y - 20
        for i, o in enumerate(obs):
            canvas.setFillColor(overdue_c if i == 0 else warn_c if i < 3 else info_c)
            canvas.setFont(hf if i < 2 else cf, 8)
            canvas.drawString(lm + 5, oy - i * 15, o)

        canvas.setFillColor(text_mut); canvas.setFont(cf, 8)
        canvas.drawString(lm, bm - 15, f"Page {doc.page}  |  Bank Submit History Report")
        canvas.drawRightString(pw - rm, bm - 15, "Confidential")
        canvas.restoreState()

    def draw_suggestions(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
        canvas.setFillColor(title_clr); canvas.setFont(hf, 20)
        canvas.drawString(lm, ph - tm - 10, "OVERDUE RECOVERY SUGGESTIONS")
        canvas.setStrokeColor(accent); canvas.setLineWidth(2)
        canvas.line(lm, ph - tm - 18, lm + 300, ph - tm - 18)
        canvas.setFillColor(text_mut); canvas.setFont(cf, 10)
        canvas.drawString(lm, ph - tm - 38, "Firm-wise & Bank-wise action plan for CO/ED review  |  " + generated_at)

        sec_y = ph - tm - 75
        canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
        canvas.drawString(lm, sec_y, "1. Firm-wise Overdue Recovery Roadmap")
        canvas.setStrokeColor(border_c); canvas.setLineWidth(0.5)
        canvas.line(lm, sec_y - 5, pw - rm, sec_y - 5)

        sug_header = ["#", "Firm Name", "Bills", "Bill Value", "Max Days", "Suggested Action", "Priority"]
        sug_rows = [sug_header]
        if len(firm_overdue) > 0:
            for rank, (firm, frow) in enumerate(firm_overdue.head(12).iterrows(), 1):
                cnt = int(frow["count"]); max_d = int(frow["max_overdue_days"])
                val = _usd(frow["total_value"])
                if max_d > 60: action = "Immediate escalation to CO. Legal notice."; priority = "URGENT"
                elif max_d > 30 and cnt > 3: action = "Schedule meeting with firm accounts dept."; priority = "HIGH"
                elif max_d > 30: action = "Send formal reminder. Follow up weekly."; priority = "MEDIUM"
                elif cnt > 5: action = "Bulk collection drive via sales team."; priority = "MEDIUM"
                else: action = "Standard follow-up via sales team."; priority = "LOW"
                sug_rows.append([str(rank), str(firm)[:30], str(cnt), val, str(max_d) + "d", action, priority])

        scol_w = [22, 150, 40, 85, 50, 310, 60]
        stbl = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
            ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
            ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (4, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
            ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
        ])
        for ri in range(1, len(sug_rows)):
            pv = sug_rows[ri][-1]
            if pv == "URGENT": stbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); stbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
            elif pv == "HIGH": stbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); stbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        st = LongTable(sug_rows, colWidths=scol_w, hAlign="LEFT")
        st.setStyle(stbl)
        stw, sth = st.wrap(uw, 200)
        st.drawOn(canvas, lm, sec_y - 25 - sth)

        sec_b_y = sec_y - 25 - sth - 30
        canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
        canvas.drawString(lm, sec_b_y, "2. Our Bank-wise Overdue Recovery Roadmap")
        canvas.setStrokeColor(border_c); canvas.line(lm, sec_b_y - 5, pw - rm, sec_b_y - 5)

        bsug_header = ["#", "Our Bank", "Bills", "Overdue Value", "Max Days", "Bank Action Required", "Priority"]
        bsug_rows = [bsug_header]
        if len(bank_overdue) > 0:
            for rank, (bname, brow) in enumerate(bank_overdue.head(10).iterrows(), 1):
                max_d = int(brow["max_overdue_days"]); cnt = int(brow["count"]); val = _usd(brow["total_value"])
                if max_d > 60: action = "Escalate to bank management. Request written update."; priority = "URGENT"
                elif max_d > 30: action = "Meet relationship manager. Push for clearance."; priority = "HIGH"
                else: action = "Regular follow-up with bank contact."; priority = "MEDIUM"
                bsug_rows.append([str(rank), str(bname)[:25], str(cnt), val, str(max_d) + "d", action, priority])

        bsc_w = [22, 100, 40, 85, 50, 355, 60]
        bstbl = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
            ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
            ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (4, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
            ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
        ])
        for ri in range(1, len(bsug_rows)):
            pv = bsug_rows[ri][-1]
            if pv == "URGENT": bstbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); bstbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
            elif pv == "HIGH": bstbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); bstbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
        bst = LongTable(bsug_rows, colWidths=bsc_w, hAlign="LEFT")
        bst.setStyle(bstbl)
        bstw, bsth = bst.wrap(uw, 200)
        bst.drawOn(canvas, lm, sec_b_y - 25 - bsth)

        box_top = sec_b_y - 25 - bsth - 25
        canvas.setFillColor(colors.HexColor("#f0f9ff")); canvas.roundRect(lm, box_top - 160, uw, 160, 8, fill=1, stroke=0)
        canvas.setStrokeColor(info_c); canvas.setLineWidth(1.5); canvas.roundRect(lm, box_top - 160, uw, 160, 8, fill=0, stroke=1)
        canvas.setFillColor(title_clr); canvas.setFont(hf, 12)
        canvas.drawString(lm + 15, box_top - 20, "3. General Recovery Strategy")
        suggestions = [
            "1. PRIORITY MATRIX: Classify overdue bills into Critical (>60d), High (>30d), Medium (>15d) and Low risk.",
            "2. FIRM ENGAGEMENT: Schedule direct meetings with top overdue firms. Prepare firm-wise outstanding statements.",
            "3. BANK COORDINATION: Work with relationship managers at CBP, SEBPLC, DBBL for faster clearance.",
            "4. SALES TEAM ALERT: Assign specific sales persons to follow up on their respective firm overdue.",
            "5. WEEKLY REVIEW: Establish weekly overdue review meeting with CO to track recovery progress.",
        ]
        sy = box_top - 38
        for i, s in enumerate(suggestions):
            canvas.setFillColor(text_dark); canvas.setFont(cf, 8)
            canvas.drawString(lm + 20, sy - i * 18, s)

        canvas.setFillColor(text_mut); canvas.setFont(cf, 8)
        canvas.drawString(lm, bm - 15, f"Page {doc.page}  |  Bank Submit History Report")
        canvas.drawRightString(pw - rm, bm - 15, "Confidential")
        canvas.restoreState()

    def draw_data_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
        canvas.setFillColor(title_clr); canvas.setFont(hf, 16)
        canvas.drawString(lm, ph - tm - 5, "DETAILED BILL DATA")
        canvas.setStrokeColor(accent); canvas.setLineWidth(1.5)
        canvas.line(lm, ph - tm - 12, lm + 180, ph - tm - 12)
        canvas.setFillColor(text_mut); canvas.setFont(cf, 9)
        canvas.drawRightString(pw - rm, ph - tm - 8, f"Page {doc.page}  |  {period_label}")
        canvas.setStrokeColor(border_c); canvas.setLineWidth(0.3)
        canvas.line(lm, bm + 5, pw - rm, bm + 5)
        canvas.setFillColor(text_mut); canvas.setFont(cf, 7)
        canvas.drawString(lm, bm - 8, f"Smart Dashboard v2.0  |  Generated: {generated_at}")
        canvas.drawRightString(pw - rm, bm - 8, "Confidential")
        canvas.restoreState()

    elements = []
    elements.append(Spacer(1, 20)); elements.append(PageBreak())
    elements.append(Spacer(1, 20)); elements.append(PageBreak())
    elements.append(Spacer(1, 20)); elements.append(PageBreak())

    dt = overdue_df.fillna("").astype(str)
    col_list = list(overdue_df.columns)
    col_widths_pdf = []
    for col in col_list:
        vals = dt[col].tolist()
        if len(vals) > 100: vals = vals[::max(1, len(vals)//100)]
        measured = [mw(v, cf, 7.5) for v in vals if v]
        mx = max([mw(col, hf, 9)] + measured) if measured else mw(col, hf, 9)
        col_widths_pdf.append(max(55, min(200, mx + 12)))
    tot_w = sum(col_widths_pdf)
    if tot_w > uw: col_widths_pdf = [w * uw / tot_w for w in col_widths_pdf]

    def chunk(cols, widths, max_w):
        groups, cur, cw3 = [], [], 0.0
        for c, w in zip(cols, widths):
            if cur and cw3 + w > max_w: groups.append(cur); cur = [c]; cw3 = w
            else: cur.append(c); cw3 += w
        if cur: groups.append(cur)
        return groups

    ss2 = getSampleStyleSheet()
    hs = ParagraphStyle("RH", parent=ss2["Normal"], fontName=hf, fontSize=8, leading=10, textColor=header_fg, alignment=TA_LEFT)
    cs = ParagraphStyle("RC", parent=ss2["Normal"], fontName=cf, fontSize=7, leading=9, alignment=TA_LEFT)

    _pg_callbacks = {2: draw_summary, 3: draw_suggestions}
    def _draw_later(canvas, doc):
        fn = _pg_callbacks.get(doc.page, draw_data_page)
        fn(canvas, doc)

    groups = chunk(col_list, col_widths_pdf, uw)
    for g_i, g_cols in enumerate(groups):
        g_w = [col_widths_pdf[col_list.index(c)] for c in g_cols]
        rows_data = [[Paragraph(str(c), hs) for c in g_cols]]
        for rv in overdue_df.fillna("").astype(str).values.tolist():
            row_cells = []
            for c in g_cols:
                idx = col_list.index(c)
                val = str(rv[idx])
                style = cs
                if c == "Due Status":
                    if "Overdue" in val:  style = ParagraphStyle("OVR"+str(g_i), parent=cs, textColor=overdue_c, fontName=hf)
                    elif "7d" in val:      style = ParagraphStyle("D7"+str(g_i), parent=cs, textColor=warn_c, fontName=hf)
                    elif "15d" in val:     style = ParagraphStyle("D15"+str(g_i), parent=cs, textColor=orange_c)
                    elif "30d" in val:     style = ParagraphStyle("D30"+str(g_i), parent=cs, textColor=info_c)
                    elif "60d" in val:     style = ParagraphStyle("D60"+str(g_i), parent=cs, textColor=safe_c)
                elif c == "Days Until Maturity":
                    try:
                        dv = int(val)
                        if dv > 0:    style = ParagraphStyle("DVp"+str(g_i), parent=cs, textColor=overdue_c, fontName=hf)
                        elif dv >= -7: style = ParagraphStyle("DV7"+str(g_i), parent=cs, textColor=warn_c, fontName=hf)
                        elif dv >= -15: style = ParagraphStyle("DV15"+str(g_i), parent=cs, textColor=orange_c)
                        elif dv >= -30: style = ParagraphStyle("DV30"+str(g_i), parent=cs, textColor=info_c)
                        else:          style = ParagraphStyle("DV60"+str(g_i), parent=cs, textColor=safe_c)
                    except: pass
                row_cells.append(Paragraph(val, style))
            rows_data.append(row_cells)

        tbl = LongTable(rows_data, repeatRows=1, colWidths=g_w, hAlign="LEFT", splitByRow=1, spaceBefore=6, spaceAfter=6)
        ts = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"), ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
            ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
        ])
        tbl.setStyle(ts)
        elements.append(tbl)
        if g_i < len(groups) - 1: elements.append(PageBreak())

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=_landscape(A3), leftMargin=lm, rightMargin=rm, topMargin=tm + 20, bottomMargin=bm)
    doc.build(elements, onFirstPage=draw_cover, onLaterPages=_draw_later)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# ADMIN PANEL
# ─────────────────────────────────────────────────────────────────────────────
def show_admin_panel():
    st.markdown("## Admin Panel")
    st.caption("User management, permissions, activity logs & security")
    st.markdown("---")

    users = load_users()
    act = load_activity()

    tab_dash, tab_users, tab_perms, tab_log, tab_sessions, tab_security, tab_email = st.tabs([
        "Dashboard", "Users", "Permissions", "Activity Log", "Sessions", "Security", "Email Schedule"
    ])

    with tab_dash:
        col1, col2, col3, col4 = st.columns(4)
        active = sum(1 for u in users.values() if u.get("status") == "active")
        disabled = len(users) - active
        total_logins = sum(u.get("login_count", 0) for u in users.values())
        col1.metric("Total Users", len(users))
        col2.metric("Active", active)
        col3.metric("Disabled", disabled)
        col4.metric("Total Logins", total_logins)

        st.markdown("#### User Login Summary")
        user_rows = []
        for uname, udata in sorted(users.items(), key=lambda x: x[1].get("login_count", 0), reverse=True):
            user_rows.append({
                "Username": uname, "Name": udata.get("name", ""),
                "Role": udata.get("role", ""), "Status": udata.get("status", ""),
                "Sales Person": udata.get("sales_person", ""),
                "Logins": udata.get("login_count", 0),
                "Last Login": udata.get("last_login", "Never"),
            })
        st.dataframe(pd.DataFrame(user_rows), use_container_width=True, hide_index=True)

        st.markdown("#### Recent Activity (Last 20)")
        if act["logs"]:
            log_df = pd.DataFrame(act["logs"][:20])
            st.dataframe(log_df, use_container_width=True, hide_index=True)
        else:
            st.info("No activity recorded yet.")

    with tab_users:
        st.markdown("### User List")

        all_user_rows = []
        for uname, udata in users.items():
            perms = udata.get("permissions", [])
            if perms == ["all"]:
                perm_str = "ALL"
            else:
                perm_str = ", ".join(perms[:5]) + ("..." if len(perms) > 5 else "")
            all_user_rows.append({
                "Username": uname,
                "Name": udata.get("name", ""),
                "Password": udata.get("password", ""),
                "Role": udata.get("role", ""),
                "Email": udata.get("email", ""),
                "Status": udata.get("status", ""),
                "Sales Person": udata.get("sales_person", ""),
                "Permissions": perm_str,
                "Created": udata.get("created_at", ""),
                "Last Login": udata.get("last_login", "Never"),
                "Login Count": udata.get("login_count", 0),
            })
        user_df = pd.DataFrame(all_user_rows)
        st.dataframe(user_df, use_container_width=True, hide_index=True, height=min(38 + len(user_df) * 35, 500))

        st.markdown("---")
        st.markdown("### Manage Users")
        col_add, col_edit = st.columns(2)

        with col_add:
            st.markdown("#### Add New User")
            new_user = st.text_input("Username", key="new_user")
            new_name = st.text_input("Full Name", key="new_name")
            new_email = st.text_input("Email", key="new_email")
            new_pass = st.text_input("Password", type="password", key="new_pass")
            new_role = st.selectbox("Role", ["admin", "manager", "sales", "custom"], key="new_role")
            if new_role == "custom":
                new_role = st.text_input("Custom Role Name", key="custom_role")
            new_sales_person = st.text_input("Sales Person Name (for data filter)", key="new_sales_person",
                                              help="If role is sales, user will only see data for this sales person. Leave blank for no filter.")
            if st.button("Add User", type="primary"):
                if new_user and new_name and new_pass:
                    uname = new_user.strip().lower()
                    if uname in users:
                        st.error(f"User '{uname}' already exists!")
                    else:
                        users[uname] = {
                            "password": new_pass, "role": new_role, "name": new_name,
                            "email": new_email, "status": "active",
                            "created_at": datetime.now().strftime("%Y-%m-%d"),
                            "last_login": "", "login_count": 0,
                            "permissions": ["all"] if new_role == "admin" else list(TAB_MAP.keys()),
                            "sales_person": new_sales_person.strip(),
                        }
                        save_users(users)
                        log_activity(st.session_state.get("username", "admin"), "user_created", f"Created user: {uname}")
                        st.success(f"User '{uname}' created!")
                        st.rerun()
                else:
                    st.warning("Fill all required fields.")

        with col_edit:
            st.markdown("#### Edit / Delete User")
            edit_user = st.selectbox("Select User", [u for u in users.keys()], key="edit_user_sel")
            if edit_user:
                udata = users[edit_user]
                ed_name = st.text_input("Name", value=udata.get("name", ""), key="ed_name")
                ed_email = st.text_input("Email", value=udata.get("email", ""), key="ed_email")
                ed_role = st.selectbox("Role", ["admin", "manager", "sales", "custom"],
                                       index=["admin","manager","sales","custom"].index(udata.get("role","sales"))
                                       if udata.get("role","sales") in ["admin","manager","sales"] else 3,
                                       key="ed_role")
                if ed_role == "custom":
                    ed_role = st.text_input("Custom Role", value=udata.get("role",""), key="ed_custom_role")
                ed_status = st.selectbox("Status", ["active", "disabled"],
                                          index=0 if udata.get("status") == "active" else 1, key="ed_status")

                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    new_pass = st.text_input("New Password (blank=no change)", type="password", key="ed_pass")
                ed_sales_person = st.text_input("Sales Person Name", value=udata.get("sales_person", ""), key="ed_sales_person",
                                                 help="User will only see data for this sales person name. Leave blank for no filter.")
                with col_s2:
                    if st.button("Update User", type="primary"):
                        users[edit_user]["name"] = ed_name
                        users[edit_user]["email"] = ed_email
                        users[edit_user]["role"] = ed_role
                        users[edit_user]["status"] = ed_status
                        users[edit_user]["sales_person"] = ed_sales_person.strip()
                        if new_pass:
                            users[edit_user]["password"] = new_pass
                        save_users(users)
                        log_activity(st.session_state.get("username", "admin"), "user_updated", f"Updated: {edit_user}")
                        st.success(f"User '{edit_user}' updated!")
                        st.rerun()

                if edit_user != "admin":
                    if st.button("Delete User", type="secondary"):
                        if st.session_state.get("confirm_delete") == edit_user:
                            del users[edit_user]
                            save_users(users)
                            log_activity(st.session_state.get("username", "admin"), "user_deleted", f"Deleted: {edit_user}")
                            st.success(f"User '{edit_user}' deleted!")
                            st.session_state.pop("confirm_delete", None)
                            st.rerun()
                        else:
                            st.session_state.confirm_delete = edit_user
                            st.warning("Click again to confirm deletion.")
                else:
                    st.info("Cannot delete admin user.")

    with tab_perms:
        st.markdown("### Permission Control")
        perm_user = st.selectbox("Select User", list(users.keys()), key="perm_user")
        if perm_user:
            udata = users[perm_user]
            current_perms = udata.get("permissions", [])
            is_all = current_perms == ["all"]
            st.markdown(f"**{udata.get('name', perm_user)}** ({udata.get('role', '')}) - Tab Access:")
            if is_all:
                st.info("Admin role - has all permissions by default. Uncheck tabs to restrict.")
            new_perms = {}
            for tab_key, tab_name in TAB_MAP.items():
                new_perms[tab_key] = st.checkbox(tab_name, value=is_all or tab_key in current_perms, key=f"perm_{perm_user}_{tab_key}")
            if st.button("Save Permissions", type="primary"):
                selected = [k for k, v in new_perms.items() if v]
                users[perm_user]["permissions"] = selected if selected else ["none"]
                save_users(users)
                log_activity(st.session_state.get("username", "admin"), "permissions_updated", f"Updated for: {perm_user}")
                st.success("Permissions saved!")
                st.rerun()

    with tab_log:
        st.markdown("### Activity Log")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            log_filter_user = st.selectbox("Filter by User", ["All"] + list(users.keys()), key="log_filter_user")
        with col_f2:
            log_filter_action = st.selectbox("Filter by Action", ["All", "login", "login_failed", "user_created", "user_updated", "user_deleted", "permissions_updated", "pdf_download"], key="log_filter_action")

        filtered_logs = act.get("logs", [])
        if log_filter_user != "All":
            filtered_logs = [l for l in filtered_logs if l.get("user") == log_filter_user]
        if log_filter_action != "All":
            filtered_logs = [l for l in filtered_logs if l.get("action") == log_filter_action]

        st.caption(f"Showing {len(filtered_logs)} of {len(act.get('logs', []))} logs")
        if filtered_logs:
            log_df = pd.DataFrame(filtered_logs)
            st.dataframe(log_df, use_container_width=True, hide_index=True, height=400)
        else:
            st.info("No logs match the filter.")

        if st.button("Clear All Logs"):
            save_activity({"logs": []})
            st.success("Activity logs cleared!")
            st.rerun()

    with tab_sessions:
        st.markdown("### Session Management")
        sessions = load_sessions()
        all_sessions = sessions.get("sessions", {})
        now = datetime.now()
        online = []
        offline = []
        for sid, sdata in all_sessions.items():
            try:
                last = datetime.strptime(sdata["last_active"], "%Y-%m-%d %H:%M:%S")
                diff = (now - last).total_seconds()
                sdata["sid"] = sid
                sdata["idle_seconds"] = int(diff)
                if diff < 300:
                    online.append(sdata)
                else:
                    offline.append(sdata)
            except Exception:
                pass

        col_s1, col_s2, col_s3 = st.columns(3)
        col_s1.metric("Online Now", len(online))
        col_s2.metric("Idle (>5min)", len(offline))
        col_s3.metric("Total Sessions", len(all_sessions))

        if online:
            st.markdown("#### Online Users")
            for s in online:
                c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
                c1.write(f"**{s.get('user', '')}**")
                c2.write(f"Login: {s.get('login_time', '')}")
                c3.write(f"Active: {s.get('last_active', '')}")
                if c4.button("Kick", key=f"kick_{s['sid']}"):
                    kill_session(s["sid"])
                    st.success(f"Kicked {s.get('user','')}")
                    st.rerun()

        if offline:
            st.markdown("#### Idle Sessions")
            for s in offline:
                c1, c2, c3 = st.columns([2, 3, 1])
                c1.write(f"**{s.get('user', '')}**")
                idle_min = s.get('idle_seconds', 0) // 60
                c2.write(f"Idle: {idle_min} min | Last: {s.get('last_active', '')}")
                if c3.button("Remove", key=f"rm_{s['sid']}"):
                    kill_session(s["sid"])
                    st.rerun()

        st.markdown("---")
        col_act1, col_act2 = st.columns(2)
        with col_act1:
            if st.button("Kick All Other Users", type="primary"):
                kill_all_sessions(exclude_user=st.session_state.get("username", ""))
                st.success("All other sessions killed!")
                st.rerun()
        with col_act2:
            if st.button("Cleanup Stale Sessions"):
                cleaned = 0
                for sid, sdata in list(all_sessions.items()):
                    try:
                        last = datetime.strptime(sdata["last_active"], "%Y-%m-%d %H:%M:%S")
                        if (now - last).total_seconds() > 3600:
                            del all_sessions[sid]
                            cleaned += 1
                    except Exception:
                        del all_sessions[sid]
                        cleaned += 1
                sessions["sessions"] = all_sessions
                save_sessions(sessions)
                st.info(f"Cleaned {cleaned} stale sessions")
                st.rerun()

        st.markdown("#### Role Distribution")
        role_counts = {}
        for u in users.values():
            r = u.get("role", "unknown")
            role_counts[r] = role_counts.get(r, 0) + 1
        fig_roles = px.pie(names=list(role_counts.keys()), values=list(role_counts.values()),
                           title="Users by Role", color_discrete_sequence=["#00c9a7", "#1a8fff", "#ff6b35", "#ffd700"])
        fig_roles.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                font=dict(color="#8899aa", size=11), margin=dict(l=8,r=8,t=32,b=8),
                                height=300)
        st.plotly_chart(fig_roles, use_container_width=True)

    with tab_security:
        st.markdown("### Security Overview")
        failed = [l for l in act.get("logs", []) if l.get("action") == "login_failed"]
        st.metric("Failed Login Attempts", len(failed))
        if failed:
            st.markdown("#### Recent Failed Attempts")
            st.dataframe(pd.DataFrame(failed[:20]), use_container_width=True, hide_index=True)
        st.markdown("#### Password Status")
        for uname, udata in users.items():
            pwd = udata.get("password", "")
            strength = "Strong" if len(pwd) >= 8 and any(c.isdigit() for c in pwd) and any(not c.isalnum() for c in pwd) else "Medium" if len(pwd) >= 6 else "Weak"
            color = "#00c9a7" if strength == "Strong" else "#ff9500" if strength == "Medium" else "#ff3b30"
            st.markdown(f"**{uname}** ({udata.get('role','')}): <span style='color:{color};font-weight:700;'>{strength}</span>", unsafe_allow_html=True)

    with tab_email:
        st.markdown("### Email Schedule — Auto Report")
        email_cfg = load_email_schedule()

        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.markdown("#### SMTP Settings")
            smtp_host = st.text_input("SMTP Host", value=email_cfg.get("smtp_host", "smtp.gmail.com"), key="smtp_host")
            smtp_port = st.number_input("SMTP Port", value=int(email_cfg.get("smtp_port", 587)), key="smtp_port")
            smtp_user = st.text_input("SMTP Username / Email", value=email_cfg.get("smtp_user", ""), key="smtp_user")
            smtp_pass = st.text_input("SMTP Password / App Password", value=email_cfg.get("smtp_pass", ""), type="password", key="smtp_pass")
            use_tls = st.checkbox("Use TLS", value=email_cfg.get("use_tls", True), key="smtp_tls")
            from_name = st.text_input("From Name", value=email_cfg.get("from_name", "Bank Dashboard Reports"), key="from_name")
            from_email = st.text_input("From Email", value=email_cfg.get("from_email", ""), key="from_email")

        with col_e2:
            st.markdown("#### Schedule Settings")
            auto_enabled = st.checkbox("Enable Auto Schedule", value=email_cfg.get("auto_enabled", False), key="auto_enabled")
            schedule = st.selectbox("Schedule Type", ["off", "daily", "weekly", "monthly"],
                                     index=["off","daily","weekly","monthly"].index(email_cfg.get("schedule", "off")),
                                     key="schedule_type")
            send_day = st.selectbox("Send Day (for weekly)",
                                     ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"],
                                     index=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"].index(email_cfg.get("send_day", "Monday")),
                                     key="send_day")
            send_time = st.text_input("Send Time (HH:MM)", value=email_cfg.get("send_time", "09:00"), key="send_time")

            st.markdown("#### Recipients")
            recipients_text = st.text_area("Email Addresses (one per line)",
                                            value="\n".join(email_cfg.get("recipients", [])),
                                            key="recipients_area",
                                            help="Enter one email address per line")
            recipients = [r.strip() for r in recipients_text.strip().split("\n") if r.strip()]

            last_sent = email_cfg.get("last_sent", "Never")
            st.info(f"Last sent: **{last_sent}**")

        if st.button("Save Email Settings", type="primary"):
            email_cfg.update({
                "smtp_host": smtp_host, "smtp_port": int(smtp_port),
                "smtp_user": smtp_user, "smtp_pass": smtp_pass,
                "use_tls": use_tls, "from_name": from_name, "from_email": from_email,
                "recipients": recipients, "schedule": schedule,
                "send_day": send_day, "send_time": send_time,
                "auto_enabled": auto_enabled,
            })
            save_email_schedule(email_cfg)
            log_activity(username, "email_config_updated", "Email schedule settings updated")
            st.success("Email settings saved!")
            st.rerun()

        st.markdown("---")
        st.markdown("#### Send Now — Manual Test")
        st.markdown(f"**Recipients:** {', '.join(recipients) if recipients else 'None'}")
        st.markdown(f"**SMTP:** {smtp_host}:{smtp_port} | **User:** {smtp_user}")

        if st.button("Send Report Now", type="primary"):
            if not recipients:
                st.error("No recipients configured!")
            elif not smtp_user:
                st.error("SMTP username not configured!")
            else:
                with st.spinner("Generating PDF and sending email..."):
                    try:
                        _fp = st.session_state.get("_email_fp", None)
                        if not _fp or not os.path.exists(_fp):
                            st.error("No data file found. Please load the dashboard first.")
                        else:
                            _raw = pd.read_excel(_fp)
                            if _raw is None or _raw.empty:
                                st.error("No data available to generate report.")
                            else:
                                from datetime import datetime as _dt
                                _today = _dt.now()
                                if "Maturity Date" in _raw.columns:
                                    _raw["Maturity Date"] = pd.to_datetime(_raw["Maturity Date"], errors="coerce")
                                if "Payment. Rcv Dt" in _raw.columns:
                                    _raw["Payment. Rcv Dt"] = pd.to_datetime(_raw["Payment. Rcv Dt"], errors="coerce")
                                _dd = _raw[_raw["Payment. Rcv Dt"].isna() & _raw["Maturity Date"].notna()].copy() if "Maturity Date" in _raw.columns and "Payment. Rcv Dt" in _raw.columns else _raw
                                if "Maturity Date" in _dd.columns:
                                    _dd["Days Until Maturity"] = (_today - _dd["Maturity Date"]).dt.days
                                    _overdue = _dd[_dd["Days Until Maturity"] > 0]
                                else:
                                    _overdue = _dd
                                if REPORTLAB_AVAILABLE:
                                    pdf_bytes = smart_overdue_pdf(_overdue, _dd, "All", _today.strftime("%d %b %Y %H:%M"))
                                else:
                                    pdf_bytes = None
                                if pdf_bytes:
                                    ok, msg = send_report_email(pdf_bytes)
                                    if ok:
                                        st.success(msg)
                                        log_activity(st.session_state.get("username", "admin"), "email_sent_manual", msg)
                                    else:
                                        st.error(msg)
                                else:
                                    st.error("Could not generate PDF")
                    except Exception as e:
                        st.error(f"Error: {str(e)}")

        st.markdown("---")
        st.markdown("#### Gmail App Password Setup")
        st.markdown("""
        1. Go to [Google Account Security](https://myaccount.google.com/security)
        2. Enable **2-Step Verification**
        3. Go to **App Passwords** → Generate one for "Mail"
        4. Use that 16-char password as SMTP Password
        5. SMTP Host: `smtp.gmail.com`, Port: `587`
        """)

# Check login
if not show_login():
    st.stop()

user_role = st.session_state.get("user_role", "admin")
user_name = st.session_state.get("user_name", "User")
username  = st.session_state.get("username", "admin")
heartbeat_session()

# ─────────────────────────────────────────────────────────────────────────────
# CHECK ADMIN PANEL MODE (after login)
# ─────────────────────────────────────────────────────────────────────────────
if user_role == "admin" and st.session_state.get("show_admin"):
    show_admin_panel()
    if st.button("Back to Dashboard"):
        st.session_state.pop("show_admin", None)
        st.rerun()
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# THEME & STYLE
# ─────────────────────────────────────────────────────────────────────────────
def _apply_theme_css(theme: str):
    if theme == "light":
        bg = "#f6f7fb"; sidebar_bg = "#ffffff"; text = "#0b1220"
        mut = "#445066"; border = "rgba(11,18,32,0.15)"
        btn_bg = "#00c9a7"; btn_fg = "#071824"; btn_bg_hover = "#00b397"
        input_bg = "#ffffff"
    else:
        bg = "#0b1220"; sidebar_bg = "#1a1f2c"; text = "#e2eaf3"
        mut = "#8899aa"; border = "rgba(226,234,243,0.16)"
        btn_bg = "#00c9a7"; btn_fg = "#071824"; btn_bg_hover = "#00b397"
        input_bg = "#0f1626"

    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    .stApp {{ background: {bg}; color: {text}; font-family: 'Inter', sans-serif; }}
    section[data-testid='stSidebar'] > div:first-child {{ background: {sidebar_bg}; }}
    textarea, input, select {{
        background: {input_bg} !important;
        color: {text} !important;
        border-color: {border} !important;
    }}
    button[kind='primary'], button[data-testid='baseButton-primary'] {{
        background: linear-gradient(135deg, {btn_bg}, #00b397) !important;
        color: {btn_fg} !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
    }}
    button[kind='primary']:hover, button[data-testid='baseButton-primary']:hover {{
        background: linear-gradient(135deg, {btn_bg_hover}, #009e83) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 15px rgba(0,201,167,0.35) !important;
    }}
    button[kind='secondary'], button[data-testid='baseButton-secondary'] {{
        background: transparent !important;
        color: {text} !important;
        border: 1px solid {border} !important;
        border-radius: 8px !important;
    }}
    [data-testid="stDialogContainer"] {{
        width: 90vw !important;
        max-width: 90vw !important;
        height: 90vh !important;
        top: 5vh !important;
        left: 5vw !important;
        padding: 0 !important;
        margin: 0 !important;
    }}
    [data-testid="stDialogContainer"] > div {{
        width: 100% !important;
        max-width: 100% !important;
        height: 100% !important;
        border-radius: 14px !important;
    }}
    [data-testid="stDialog"] {{
        width: 100% !important;
        max-width: 100% !important;
        height: 100% !important;
        border-radius: 14px !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    [data-testid="stDialog"] > div {{
        width: 100% !important;
        max-width: 100% !important;
        height: 100% !important;
        border-radius: 14px !important;
    }}
    section[data-testid="stMain"] {{
        max-width: 100% !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }}
    .metric-card {{
        background: linear-gradient(145deg, rgba(0,201,167,0.12), rgba(0,201,167,0.03));
        border: 1px solid rgba(0,201,167,0.25);
        border-radius: 16px;
        padding: 22px 24px;
        box-shadow: 0 6px 20px rgba(0,201,167,0.12), 0 2px 6px rgba(0,0,0,0.2), inset 0 1px 0 rgba(255,255,255,0.06);
        transform: perspective(900px) rotateX(1deg) rotateY(0deg) translateZ(0);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }}
    .metric-card::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,201,167,0.5), transparent);
    }}
    .metric-card::after {{
        content: '';
        position: absolute;
        top: 0; left: 0; bottom: 0;
        width: 1px;
        background: linear-gradient(180deg, rgba(0,201,167,0.5), transparent);
    }}
    .metric-card:hover {{
        transform: perspective(900px) rotateX(0deg) rotateY(-1deg) translateZ(6px);
        box-shadow: 0 12px 35px rgba(0,201,167,0.2), 0 4px 10px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.1);
        border-color: rgba(0,201,167,0.45);
    }}
    .metric-label {{ font-size: 13px; color: #8899aa; font-weight: 600; letter-spacing: 0.03em; margin-bottom: 6px; }}
    .metric-value {{ font-size: 28px; font-weight: 800; color: #00c9a7; line-height: 1.1; }}
    .metric-delta {{ font-size: 12px; color: #00c9a7; margin-top: 6px; font-weight: 500; }}
    .alert-critical {{
        background: linear-gradient(145deg, rgba(255,59,48,0.18), rgba(255,59,48,0.06));
        border: 1px solid rgba(255,59,48,0.5);
        border-left: 4px solid #ff3b30;
        border-radius: 14px;
        padding: 18px 22px;
        margin: 8px 0;
        box-shadow: 0 4px 15px rgba(255,59,48,0.15), 0 1px 3px rgba(0,0,0,0.2), inset 0 1px 0 rgba(255,255,255,0.05);
        transform: perspective(800px) rotateY(0deg) translateZ(0);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }}
    .alert-critical::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,59,48,0.6), transparent);
    }}
    .alert-critical:hover {{
        transform: perspective(800px) rotateY(-1deg) translateZ(4px);
        box-shadow: 0 8px 25px rgba(255,59,48,0.25), 0 2px 6px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.08);
    }}
    .alert-warning {{
        background: linear-gradient(145deg, rgba(255,149,0,0.18), rgba(255,149,0,0.06));
        border: 1px solid rgba(255,149,0,0.5);
        border-left: 4px solid #ff9500;
        border-radius: 14px;
        padding: 18px 22px;
        margin: 8px 0;
        box-shadow: 0 4px 15px rgba(255,149,0,0.12), 0 1px 3px rgba(0,0,0,0.2), inset 0 1px 0 rgba(255,255,255,0.05);
        transform: perspective(800px) rotateY(0deg) translateZ(0);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }}
    .alert-warning::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,149,0,0.6), transparent);
    }}
    .alert-warning:hover {{
        transform: perspective(800px) rotateY(-1deg) translateZ(4px);
        box-shadow: 0 8px 25px rgba(255,149,0,0.2), 0 2px 6px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.08);
    }}
    .alert-info {{
        background: linear-gradient(145deg, rgba(0,122,255,0.15), rgba(0,122,255,0.05));
        border: 1px solid rgba(0,122,255,0.45);
        border-left: 4px solid #007aff;
        border-radius: 14px;
        padding: 18px 22px;
        margin: 8px 0;
        box-shadow: 0 4px 15px rgba(0,122,255,0.12), 0 1px 3px rgba(0,0,0,0.2), inset 0 1px 0 rgba(255,255,255,0.05);
        transform: perspective(800px) rotateY(0deg) translateZ(0);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }}
    .alert-info::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,122,255,0.6), transparent);
    }}
    .alert-info:hover {{
        transform: perspective(800px) rotateY(-1deg) translateZ(4px);
        box-shadow: 0 8px 25px rgba(0,122,255,0.2), 0 2px 6px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.08);
    }}
    .alert-title {{ font-weight: 700; font-size: 15px; margin-bottom: 4px; }}
    .alert-msg {{ font-size: 13px; color: #aabbcc; line-height: 1.5; }}
    .leaderboard-card {{
        background: linear-gradient(135deg, rgba(0,201,167,0.08), rgba(26,143,255,0.06));
        border: 1px solid rgba(0,201,167,0.2);
        border-radius: 14px;
        padding: 14px 18px;
        margin: 6px 0;
        transition: all 0.2s ease;
    }}
    .leaderboard-card:hover {{
        transform: translateX(4px);
        border-color: rgba(0,201,167,0.5);
    }}
    p.sh {{ color:#8899aa; font-size:13px; font-weight:600; margin-bottom:4px; letter-spacing:.05em; }}
    .stTabs [data-baseweb="tab"] {{ font-weight: 500; }}
    .stTabs [aria-selected="true"] {{ color: #00c9a7 !important; border-bottom-color: #00c9a7 !important; }}
    [data-testid="stExpander"] {{ max-width: 100% !important; }}
    [data-testid="stExpander"] > details {{ width: 100% !important; }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
user_role = st.session_state.get("user_role", "admin")
user_name = st.session_state.get("user_name", "User")
username  = st.session_state.get("username", "admin")

_sp_filter = ""
_user_udata = load_users().get(username, {})
if user_role == "sales" and _user_udata.get("sales_person"):
    _sp_filter = _user_udata["sales_person"]

sp_badge = f"<div style='font-size:10px; color:#ff9500; margin-top:4px;'>Filtered: {_sp_filter}</div>" if _sp_filter else ""
st.sidebar.markdown(f"""
<div style='text-align:center; padding:12px 0 8px;'>
    <div style='font-size:36px;'>🏦</div>
    <div style='font-size:16px; font-weight:700; color:#00c9a7; margin:4px 0;'>Bank Submit</div>
    <div style='font-size:11px; color:#556677;'>Dashboard v2.0</div>
    <div style='background:rgba(0,201,167,0.12); border-radius:20px; padding:4px 12px; margin:8px auto; width:fit-content; font-size:11px; color:#00c9a7;'>
        👤 {user_name} <span style='color:#445566;'>({user_role})</span>
    </div>
    {sp_badge}
</div>
""", unsafe_allow_html=True)

theme_choice = st.sidebar.radio("🎨 Theme", ["dark", "light"], index=0, horizontal=True)
_apply_theme_css(theme_choice)

if user_role == "admin":
    if st.sidebar.button("⚙️ Admin Panel", use_container_width=True):
        st.session_state.show_admin = True
        st.rerun()

if st.sidebar.button("🚪 Logout", use_container_width=True):
    for k in ["authenticated", "username", "user_role", "user_name"]:
        st.session_state.pop(k, None)
    st.query_params.clear()
    st.rerun()

st.sidebar.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# PLOTLY THEME
# ─────────────────────────────────────────────────────────────────────────────
PL = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#8899aa", size=11, family="Inter, monospace"),
    margin=dict(l=8, r=8, t=32, b=8),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=10)),
    xaxis=dict(gridcolor="#1a2a3a", linecolor="#1a2a3a"),
    yaxis=dict(gridcolor="#1a2a3a", linecolor="#1a2a3a"),
)
PL_GENERAL = {k: v for k, v in PL.items() if k not in ("legend", "xaxis", "yaxis")}

C = ["#00c9a7","#1a8fff","#ff6b35","#ffd700","#cc44ff",
     "#ff3366","#44aaff","#ff9944","#66dd66","#00aaff"]

REQUIRED_COLUMNS = [
    "Firm Name", "Sales Person", "Bank Submition Date", "Invoice Value",
    "Lc Value", "Maturity Date", "Payment. Rcv Dt", "Bank Accept Date",
    "LC No", "Our Bank", "Party Name", "Bank Name"
]

def usd(v):
    try: v = float(v)
    except Exception: return "$0.00"
    if v >= 1e6: return f"${v/1e6:.2f}M"
    if v >= 1e3: return f"${v/1e3:.1f}K"
    return f"${v:.2f}"

def sh(label):
    st.markdown(f'<p class="sh">{label}</p>', unsafe_allow_html=True)

def norm_tenor(t):
    if pd.isna(t): return "Unknown"
    tt = str(t).strip()
    if tt.startswith("120"): return "120 Days"
    if tt.startswith("90"):  return "90 Days"
    if tt == "0" or "at sight" in tt.lower(): return "At Sight"
    return tt

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADER
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="⏳ Loading data…")
def load(path):
    xls = pd.ExcelFile(path)
    sheet_candidates = ["Raw Data", "raw data", "Sheet1", "Sheet 1"]
    selected = None
    for s in sheet_candidates:
        if s in xls.sheet_names:
            selected = s; break
    if selected is None:
        for s in xls.sheet_names:
            if "bank" in s.lower() and "history" in s.lower():
                selected = s; break
    if selected is None:
        selected = xls.sheet_names[0]

    df = pd.read_excel(path, sheet_name=selected)
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet '{selected}' — Missing columns: {', '.join(missing)}")

    df = df.dropna(subset=["Firm Name"])
    df["Sales Person"] = (df["Sales Person"].astype(str).str.strip()
                          .str.replace("_x000D_\n", "", regex=False).str.strip())
    df["Sales Person"] = df["Sales Person"].replace(r"^(nan|NaN|none|None|\s*)$", None, regex=True)
    df.loc[df["Sales Person"].isin([None, ""]), "Sales Person"] = None

    df["_date"]     = pd.to_datetime(df["Bank Submition Date"], errors="coerce")
    df["MonthSort"] = df["_date"].dt.to_period("M")
    df["Month"]     = df["_date"].dt.strftime("%b %Y")
    df["WeekSort"]  = df["_date"].dt.to_period("W")
    df["Week"]      = df["_date"].dt.strftime("W%V") + " " + df["_date"].dt.strftime("%b %Y")
    df["DayName"]   = df["_date"].dt.strftime("%a")
    df["Date"]      = df["_date"].dt.strftime("%d %b %Y")

    # parse date columns for due-date tracker
    for col in ["Maturity Date", "Payment. Rcv Dt", "Bank Accept Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

# ─────────────────────────────────────────────────────────────────────────────
# FILE PICKER
# ─────────────────────────────────────────────────────────────────────────────
xlsx = [f for f in glob.glob("*.xlsx") + glob.glob("**/*.xlsx", recursive=True)
        if "Dashboard" not in f]
up = st.sidebar.file_uploader("📂 Upload Excel file", type=["xlsx"])
if up:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    shutil.copyfileobj(up, tmp); tmp.close(); FP = tmp.name
    try:
        import atexit
        atexit.register(os.unlink, FP)
    except Exception:
        pass
elif xlsx:
    FP = st.sidebar.selectbox("Or pick a file", xlsx)
else:
    st.warning("⚠️ Please upload your Excel file using the sidebar."); st.stop()

raw = load(FP)
st.session_state["_email_fp"] = FP

# ── Sales person row-level security ──────────────────────────────────────────
_user_data = load_users().get(username, {})
if user_role == "sales" and _user_data.get("sales_person"):
    sp_filter = _user_data["sales_person"].strip()
    if sp_filter:
        raw = raw[raw["Sales Person"].astype(str).str.strip() == sp_filter]
        if raw.empty:
            st.warning(f"No data found for Sales Person: {sp_filter}")
            st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# SMART SEARCH (sidebar)
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 Smart Search")
smart_query = st.sidebar.text_input(
    "Search (e.g. 'DBBL June unpaid')",
    placeholder="Type bank, firm, LC#, month…",
    key="smart_search"
)

def apply_smart_search(df_in, query):
    """Parse natural language query and filter dataframe."""
    if not query or not query.strip():
        return df_in
    q = query.lower().strip()
    result = df_in.copy()

    # LC number direct match
    lc_match = re.search(r'\blc[\s#:]*([A-Za-z0-9/\-]+)', q)
    if lc_match:
        lc_val = lc_match.group(1).upper()
        result = result[result["LC No"].astype(str).str.contains(lc_val, case=False, na=False)]
        return result

    # Payment status keywords
    if any(w in q for w in ["unpaid", "not paid", "pending", "outstanding"]):
        result = result[result["Payment. Rcv Dt"].isna()]
    elif any(w in q for w in ["paid", "payment received"]):
        result = result[result["Payment. Rcv Dt"].notna()]

    if any(w in q for w in ["overdue", "matured"]):
        today = pd.Timestamp.today().normalize()
        result = result[result["Maturity Date"].notna() & (result["Maturity Date"] < today) & result["Payment. Rcv Dt"].isna()]

    if any(w in q for w in ["accepted", "bank accept"]):
        result = result[result["Bank Accept Date"].notna()]

    if any(w in q for w in ["not accepted", "unaccepted"]):
        result = result[result["Bank Accept Date"].isna()]

    # Month keywords
    months_map = {
        "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
        "may": "May", "jun": "Jun", "jul": "Jul", "aug": "Aug",
        "sep": "Sep", "oct": "Oct", "nov": "Nov", "dec": "Dec"
    }
    for abbr, full in months_map.items():
        if abbr in q:
            result = result[result["Month"].str.contains(full, case=False, na=False)]
            break

    # Week keywords
    if "last week" in q:
        today = pd.Timestamp.today()
        week_start = (today - timedelta(days=today.weekday() + 7)).normalize()
        week_end = week_start + timedelta(days=6)
        result = result[(result["_date"] >= week_start) & (result["_date"] <= week_end)]
    elif "this week" in q:
        today = pd.Timestamp.today()
        week_start = (today - timedelta(days=today.weekday())).normalize()
        result = result[result["_date"] >= week_start]

    # Bank name
    for col in ["Our Bank", "Bank Name"]:
        if col in result.columns:
            # look for known bank keywords in query
            for bank in result[col].dropna().unique():
                if str(bank).lower() in q:
                    result = result[result[col].str.lower() == str(bank).lower()]
                    break

    # Firm name
    for firm in result["Firm Name"].dropna().unique():
        if str(firm).lower() in q:
            result = result[result["Firm Name"].str.lower() == str(firm).lower()]
            break

    # Sales person
    for sp in result["Sales Person"].dropna().unique():
        if str(sp).lower() in q:
            result = result[result["Sales Person"].str.lower() == str(sp).lower()]
            break

    # Fallback: general text search across all columns
    if len(result) == len(df_in) and query.strip():
        mask = df_in.astype(str).apply(
            lambda c: c.str.contains(query.strip(), case=False, na=False)
        ).any(axis=1)
        result = df_in[mask]

    return result

# ─────────────────────────────────────────────────────────────────────────────
# FILTERS
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.markdown("### 🔽 Filters")
ml  = [str(m) for m in sorted(raw["MonthSort"].dropna().unique())]
sm  = st.sidebar.multiselect("Month",    ml, default=ml)
sf  = st.sidebar.multiselect("Firm",     sorted(raw["Firm Name"].dropna().unique()),
                              default=sorted(raw["Firm Name"].dropna().unique()))
sb  = st.sidebar.multiselect("Our Bank", sorted(raw["Our Bank"].dropna().unique()),
                              default=sorted(raw["Our Bank"].dropna().unique()))

sales_persons = sorted(raw["Sales Person"].dropna().unique())
ss_choices    = (["(Blank)"] + sales_persons) if raw["Sales Person"].isna().any() else sales_persons
ss = st.sidebar.multiselect("Sales Person", ss_choices, default=ss_choices)

sparty = st.sidebar.multiselect("Party Name",
    sorted(raw["Party Name"].dropna().unique()),
    default=sorted(raw["Party Name"].dropna().unique()))

min_date = raw["_date"].min(); max_date = raw["_date"].max()
if pd.isna(min_date) or pd.isna(max_date):
    date_range = st.sidebar.date_input("Date Range",
        value=(pd.Timestamp.today().date(), pd.Timestamp.today().date()))
else:
    date_range = st.sidebar.date_input("Date Range",
        value=(min_date.date(), max_date.date()))

df = raw.copy()
if sm:     df = df[df["MonthSort"].astype(str).isin(sm)]
if sf:     df = df[df["Firm Name"].isin(sf)]
if sb:     df = df[df["Our Bank"].isin(sb)]
if ss:
    if "(Blank)" in ss:
        sel = [s for s in ss if s != "(Blank)"]
        df = df[(df["Sales Person"].isin(sel)) | df["Sales Person"].isna()]
    else:
        df = df[df["Sales Person"].isin(ss)]
if sparty: df = df[df["Party Name"].isin(sparty)]
if isinstance(date_range, tuple) and len(date_range) == 2:
    s_d, e_d = date_range
    df = df[(df["_date"].dt.date >= s_d) & (df["_date"].dt.date <= e_d)]

# Apply smart search
if smart_query:
    df = apply_smart_search(df, smart_query)

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(df):,}** of **{len(raw):,}** records")
if smart_query:
    st.sidebar.caption(f"🔍 Smart search: *{smart_query}*")
if df.empty: st.warning("No records match the filters."); st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# AGGREGATES
# ─────────────────────────────────────────────────────────────────────────────
N      = len(df)
inv    = df["Invoice Value"].sum()
lc     = df["Lc Value"].sum()
mat_v  = df[df["Maturity Date"].notna()]["Invoice Value"].sum()
mat_n  = int(df["Maturity Date"].notna().sum())
pay_v  = df[df["Payment. Rcv Dt"].notna()]["Invoice Value"].sum()
paid_n = int(df["Payment. Rcv Dt"].notna().sum())
acc_n  = int((df["Bank Accept Date"].notna() & df["Payment. Rcv Dt"].isna()).sum())
nacc_n = int(df["Bank Accept Date"].isna().sum())
acc_v  = df[df["Bank Accept Date"].notna() & df["Payment. Rcv Dt"].isna()]["Invoice Value"].sum()
nacc_v = df[df["Bank Accept Date"].isna()]["Invoice Value"].sum()

monthly = (df.groupby(["MonthSort","Month"])
             .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"), LC=("Lc Value","sum"))
             .reset_index().sort_values("MonthSort"))
by_firm = (df.groupby("Firm Name")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False))
by_bank = (df.groupby("Our Bank")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False))
t_party = (df.groupby("Party Name")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False).head(10))
t_bname = (df.groupby("Bank Name")
             .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
             .reset_index().sort_values("Inv", ascending=False).head(10))

spg   = (df[df["Sales Person"].notna()]
          .groupby("Sales Person")
          .agg(Inv=("Invoice Value","sum"), N=("LC No","count"))
          .reset_index().sort_values("Inv", ascending=False))
sp_p  = (df[df["Payment. Rcv Dt"].notna() & df["Sales Person"].notna()]
          .groupby("Sales Person").size().reset_index(name="Paid"))
spg   = spg.merge(sp_p, on="Sales Person", how="left").fillna(0)
spg["Pct"] = (spg["Paid"] / spg["N"] * 100).round(1)

# Weekly
weekly = (df.groupby(["WeekSort","Week"])
           .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"), LC=("Lc Value","sum"),
                Paid_n=("Payment. Rcv Dt", lambda x: x.notna().sum()))
           .reset_index().sort_values("WeekSort"))
weekly["Paid_pct"] = (weekly["Paid_n"] / weekly["Count"] * 100).round(1)

wk_firm = (df.groupby(["WeekSort","Week","Firm Name"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))
wk_sp   = (df[df["Sales Person"].notna()]
             .groupby(["WeekSort","Week","Sales Person"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))
wk_bank = (df.groupby(["WeekSort","Week","Our Bank"])
             .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
             .reset_index().sort_values("WeekSort"))

wk_status = df.copy()
wk_status["Status"] = wk_status.apply(lambda r:
    "Paid"         if pd.notna(r["Payment. Rcv Dt"]) else
    "Accepted"     if pd.notna(r["Bank Accept Date"]) else
    "Not Accepted", axis=1)
wk_st_grp = (wk_status.groupby(["WeekSort","Week","Status"])
              .agg(Count=("LC No","count"), Inv=("Invoice Value","sum"))
              .reset_index().sort_values("WeekSort"))

wk_party     = (df.groupby(["WeekSort","Week","Party Name"])
                  .agg(Inv=("Invoice Value","sum"), Count=("LC No","count"))
                  .reset_index().sort_values("WeekSort"))
wk_party_top = wk_party[wk_party["Party Name"].isin(t_party["Party Name"].tolist())]

period = (f"{monthly['Month'].iloc[0]} – {monthly['Month'].iloc[-1]}"
          if len(monthly) else "—")

# ─────────────────────────────────────────────────────────────────────────────
# ALERTS ENGINE
# ─────────────────────────────────────────────────────────────────────────────
today = pd.Timestamp.today().normalize()

def compute_alerts(df_in):
    alerts = []
    # Overdue: maturity passed, not paid
    if "Maturity Date" in df_in.columns:
        overdue = df_in[
            df_in["Maturity Date"].notna() &
            (df_in["Maturity Date"] < today) &
            df_in["Payment. Rcv Dt"].isna()
        ]
        if len(overdue) > 0:
            total_overdue_val = overdue["Invoice Value"].sum()
            alerts.append({
                "level": "critical",
                "icon": "🔴",
                "title": f"Overdue Payments — {len(overdue)} Bill(s)",
                "msg": f"Maturity passed but payment not received. Total: **{usd(total_overdue_val)}**",
                "df": overdue
            })

    # Due soon: maturity in next 7 days
    if "Maturity Date" in df_in.columns:
        due_7 = df_in[
            df_in["Maturity Date"].notna() &
            (df_in["Maturity Date"] >= today) &
            (df_in["Maturity Date"] <= today + timedelta(days=7)) &
            df_in["Payment. Rcv Dt"].isna()
        ]
        if len(due_7) > 0:
            alerts.append({
                "level": "warning",
                "icon": "🟡",
                "title": f"Due in 7 Days — {len(due_7)} Bill(s)",
                "msg": f"These Bills mature within 7 days. Total: **{usd(due_7['Invoice Value'].sum())}**",
                "df": due_7
            })

    # Due in 15 days
    if "Maturity Date" in df_in.columns:
        due_15 = df_in[
            df_in["Maturity Date"].notna() &
            (df_in["Maturity Date"] > today + timedelta(days=7)) &
            (df_in["Maturity Date"] <= today + timedelta(days=15)) &
            df_in["Payment. Rcv Dt"].isna()
        ]
        if len(due_15) > 0:
            alerts.append({
                "level": "info",
                "icon": "🔵",
                "title": f"Due in 8-15 Days — {len(due_15)} Bill(s)",
                "msg": f"Upcoming maturities. Total: **{usd(due_15['Invoice Value'].sum())}**",
                "df": due_15
            })

    # Anomaly: firm with submission > 50% drop vs previous month
    if len(monthly) >= 2:
        last_m  = str(monthly["MonthSort"].iloc[-1])
        prev_m  = str(monthly["MonthSort"].iloc[-2])
        last_df = df_in[df_in["MonthSort"].astype(str) == last_m]
        prev_df = df_in[df_in["MonthSort"].astype(str) == prev_m]
        last_by_firm = last_df.groupby("Firm Name").size()
        prev_by_firm = prev_df.groupby("Firm Name").size()
        for firm in prev_by_firm.index:
            prev_cnt = prev_by_firm[firm]
            last_cnt = last_by_firm.get(firm, 0)
            if prev_cnt >= 5 and last_cnt < prev_cnt * 0.5:
                drop_pct = (1 - last_cnt/prev_cnt)*100
                alerts.append({
                    "level": "warning",
                    "icon": "🟠",
                    "title": f"Submission Drop — {firm}",
                    "msg": f"**{drop_pct:.0f}%** drop vs previous month ({int(prev_cnt)} → {int(last_cnt)} submissions)",
                    "df": None
                })

    return alerts

alerts = compute_alerts(df)

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
h1, h2 = st.columns([5, 1])
with h1:
    st.markdown("# 🏦 Bank Submit History Dashboard")
    st.markdown(
        f"<p style='color:#556677;font-size:12px;margin-top:-12px;'>"
        f"Period: <b style='color:#00c9a7'>{period}</b> &nbsp;|&nbsp; "
        f"File: {os.path.basename(FP)}</p>", unsafe_allow_html=True)
with h2:
    if alerts:
        crit = sum(1 for a in alerts if a["level"] == "critical")
        warn = sum(1 for a in alerts if a["level"] == "warning")
        st.markdown(f"""
        <div style='text-align:right; padding-top:10px;'>
            <span style='background:rgba(255,59,48,0.15); border:1px solid #ff3b30; border-radius:20px;
                         padding:4px 12px; font-size:12px; color:#ff3b30; font-weight:600;'>
                🔴 {crit} Critical
            </span>
            <br><span style='background:rgba(255,149,0,0.12); border:1px solid #ff9500; border-radius:20px;
                         padding:4px 12px; font-size:12px; color:#ff9500; font-weight:600; margin-top:4px; display:inline-block;'>
                🟡 {warn} Warnings
            </span>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# ALERTS PANEL
# ─────────────────────────────────────────────────────────────────────────────
if alerts:
    st.markdown(f"⚠️ **Smart Alerts** — {len(alerts)} active alerts (click to expand)")
    for ai, a in enumerate(alerts):
        css_cls = {"critical": "alert-critical", "warning": "alert-warning", "info": "alert-info"}.get(a["level"], "alert-info")
        alert_df = a.get("df")
        alert_count = len(alert_df) if alert_df is not None else 0
        alert_val = alert_df["Invoice Value"].sum() if alert_df is not None and "Invoice Value" in alert_df.columns else 0
        with st.expander(f"{a['icon']} {a['title']}", expanded=False):
            st.markdown(f"<div class='{css_cls}' style='margin:-12px -12px 12px -12px;'><div class='alert-msg'>{a['msg']}</div></div>", unsafe_allow_html=True)
            if alert_df is not None and len(alert_df) > 0:
                pc1, pc2, pc3 = st.columns(3)
                pc1.metric("Total Bills", f"{alert_count:,}")
                pc2.metric("Total Value", usd(alert_val))
                pc3.metric("Average Value", usd(alert_val / alert_count) if alert_count > 0 else usd(0))
                if "Firm Name" in alert_df.columns and "Invoice Value" in alert_df.columns:
                    firm_agg = alert_df.groupby("Firm Name")["Invoice Value"].sum().sort_values(ascending=True).tail(15)
                    if len(firm_agg) > 0:
                        fig_alert = go.Figure(go.Bar(
                            x=firm_agg.values, y=firm_agg.index, orientation="h",
                            marker=dict(color=firm_agg.values, colorscale="Tealgrn"),
                            text=[usd(v) for v in firm_agg.values], textposition="outside"
                        ))
                        fig_alert.update_layout(
                            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                            font=dict(color="#8899aa", size=11), margin=dict(l=8, r=8, t=32, b=8),
                            height=max(280, len(firm_agg) * 30),
                            title=dict(text=f"Top Firms — {a['title']}", font=dict(color="#00c9a7", size=14)),
                            xaxis=dict(gridcolor="#1a2a3a", title="Invoice Value (USD)"),
                            yaxis=dict(gridcolor="#1a2a3a"),
                        )
                        st.plotly_chart(fig_alert, use_container_width=True)
                show_cols = ["Firm Name", "LC No", "Bank Refno", "Party Name", "Our Bank", "Bank Name", "Maturity Date", "Days Until Maturity", "Invoice Value", "Sales Person", "Due Status"]
                show_cols = [c for c in show_cols if c in alert_df.columns or c == "Days Until Maturity"]
                alert_work = alert_df.copy()
                if "Maturity Date" in alert_work.columns:
                    mat_dt = pd.to_datetime(alert_work["Maturity Date"], errors="coerce")
                    alert_work["Days Until Maturity"] = (pd.Timestamp.today().normalize() - mat_dt).dt.days
                    alert_work["Due Status"] = alert_work["Days Until Maturity"].apply(
                        lambda d: "🔴 Overdue" if d > 0
                        else "🟡 Due in 7d" if d >= -7
                        else "🟠 Due in 15d" if d >= -15
                        else "🔵 Due in 30d" if d >= -30
                        else "🟢 Due in 60d+" if d >= -60
                        else "✅ Safe"
                    )
                pop_df = alert_work[show_cols].copy()
                pop_df = pop_df.sort_values("Days Until Maturity", ascending=False)
                for dc in ["Maturity Date", "Payment. Rcv Dt", "Bank Accept Date", "Bank Submition Date"]:
                    if dc in pop_df.columns:
                        pop_df[dc] = pop_df[dc].dt.strftime("%d %b %Y")
                if "Invoice Value" in pop_df.columns:
                    pop_df["Invoice Value"] = pop_df["Invoice Value"].map(lambda x: f"${x:,.2f}")
                if "Days Until Maturity" in pop_df.columns:
                    pop_df["Days Until Maturity"] = pop_df["Days Until Maturity"].map(lambda x: f"{int(x)}" if pd.notna(x) else "")
                st.dataframe(pop_df, use_container_width=True, hide_index=True, height=min(650, 38 * len(pop_df) + 38))
    st.markdown("")

# KPI METRICS
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"""<div class="metric-card"><div class="metric-label">📋 Total Submissions</div><div class="metric-value">{N:,}</div><div class="metric-delta">{N} total records</div></div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="metric-card"><div class="metric-label">💵 Total Invoice Value</div><div class="metric-value">{usd(inv)}</div><div class="metric-delta">{usd(inv)} total invoiced</div></div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="metric-card"><div class="metric-label">📅 Maturity Received Value</div><div class="metric-value">{usd(mat_v)}</div><div class="metric-delta">{mat_n} records · {mat_n/N*100:.1f}%</div></div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""<div class="metric-card"><div class="metric-label">✅ Payment Received Value</div><div class="metric-value">{usd(pay_v)}</div><div class="metric-delta">{paid_n} records · {paid_n/N*100:.1f}%</div></div>""", unsafe_allow_html=True)
st.markdown("")

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
(t_daily, t_overview, t_weekly, t_firm, t_banks, t_parties,
 t_payment, t_accept, t_asm, t_due, t_leaderboard) = st.tabs([
    "📅 Daily Analysis",
    "📊 Overview",
    "📅 Weekly Analysis",
    "🏢 Firm & Sales Person",
    "🏦 Banks",
    "👥 Top Parties",
    "🔄 Payment Status",
    "✅ Bank Accept Analysis",
    "📊 Asm Analysis",
    "🗓️ Due Date Tracker",
    "🏆 Leaderboard",
])

# ─────────────────────────────────────────────────────────────────────────────
# FIRM DETAIL DIALOG — popup modal for Due Date Tracker
# ─────────────────────────────────────────────────────────────────────────────
@st.dialog("📋 Firm Overdue Detail", width="large")
def show_firm_detail(firm_name):
    _overdue = st.session_state.get("_firm_overdue_data")
    if _overdue is None or firm_name not in _overdue["Firm Name"].values:
        st.error("No data found for this firm.")
        return

    firm_records = _overdue[_overdue["Firm Name"] == firm_name].copy()
    firm_total = firm_records["Invoice Value"].sum()
    firm_count = len(firm_records)
    firm_max = firm_records["Days Until Maturity"].max()
    firm_risk = "CRITICAL" if firm_max > 60 else "HIGH" if firm_max > 30 else "MEDIUM" if firm_max > 14 else "LOW"
    firm_avg = firm_records["Days Until Maturity"].mean()
    risk_color = {"CRITICAL":"#ff3b30","HIGH":"#ff9500","MEDIUM":"#ff6b35","LOW":"#00c9a7"}.get(firm_risk,"#00c9a7")

    st.markdown(f"""
    <div style="background: linear-gradient(145deg, rgba(0,201,167,0.10), rgba(26,143,255,0.06));
                border: 1px solid rgba(0,201,167,0.25); border-radius:18px;
                padding:24px 28px; margin:12px 0;
                box-shadow: 0 8px 30px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.06);
                transform: perspective(800px) rotateX(1deg);
                position:relative; overflow:hidden;">
        <div style="position:absolute; top:0; left:0; right:0; height:1px;
                    background: linear-gradient(90deg, transparent, rgba(0,201,167,0.6), transparent);"></div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
            <div>
                <div style="font-size:22px; font-weight:800; color:#e2eaf3;">{firm_name}</div>
                <div style="font-size:13px; color:#8899aa; margin-top:4px;">Full overdue bill breakdown</div>
            </div>
            <div style="background:{risk_color}22; color:{risk_color}; border:1px solid {risk_color}55;
                        border-radius:10px; padding:6px 16px; font-weight:700; font-size:14px;">{firm_risk}</div>
        </div>
        <div style="display:flex; gap:16px; flex-wrap:wrap;">
            <div style="flex:1; min-width:130px; background:rgba(0,0,0,0.2); border-radius:12px; padding:14px; text-align:center;">
                <div style="font-size:11px; color:#8899aa; text-transform:uppercase; letter-spacing:0.5px;">Total Bills</div>
                <div style="font-size:24px; font-weight:800; color:#00c9a7;">{firm_count}</div>
            </div>
            <div style="flex:1; min-width:130px; background:rgba(0,0,0,0.2); border-radius:12px; padding:14px; text-align:center;">
                <div style="font-size:11px; color:#8899aa; text-transform:uppercase; letter-spacing:0.5px;">Total Value</div>
                <div style="font-size:24px; font-weight:800; color:#00c9a7;">{usd(firm_total)}</div>
            </div>
            <div style="flex:1; min-width:130px; background:rgba(0,0,0,0.2); border-radius:12px; padding:14px; text-align:center;">
                <div style="font-size:11px; color:#8899aa; text-transform:uppercase; letter-spacing:0.5px;">Max Overdue</div>
                <div style="font-size:24px; font-weight:800; color:{risk_color};">{int(firm_max)}d</div>
            </div>
            <div style="flex:1; min-width:130px; background:rgba(0,0,0,0.2); border-radius:12px; padding:14px; text-align:center;">
                <div style="font-size:11px; color:#8899aa; text-transform:uppercase; letter-spacing:0.5px;">Avg Overdue</div>
                <div style="font-size:24px; font-weight:800; color:#1a8fff;">{firm_avg:.0f}d</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    firm_bank = firm_records.groupby("Our Bank").agg(
        Value=("Invoice Value","sum"), Count=("Invoice Value","size")
    ).sort_values("Value", ascending=False).reset_index()
    fig_fb = go.Figure(go.Bar(
        x=firm_bank["Our Bank"], y=firm_bank["Value"],
        marker=dict(color=firm_bank["Value"], colorscale="Tealgrn"),
        text=[f"{usd(v)} ({int(c)})" for v, c in zip(firm_bank["Value"], firm_bank["Count"])],
        textposition="outside"))
    fig_fb.update_layout(**PL_GENERAL, height=280,
        title=dict(text=f"Bank Distribution — {firm_name}", font=dict(color="#00c9a7", size=14)),
        xaxis=dict(gridcolor="#1a2a3a"), yaxis=dict(title="Value (USD)", gridcolor="#1a2a3a"),
        showlegend=False)
    st.plotly_chart(fig_fb, use_container_width=True)

    detail_cols = ["LC No","Bank Refno","Party Name","Our Bank","Bank Name","Maturity Date",
                   "Invoice Value","Days Until Maturity","Due Status","Sales Person"]
    detail_cols = [c for c in detail_cols if c in firm_records.columns]
    firm_detail = firm_records[detail_cols].copy()
    firm_detail = firm_detail.sort_values("Days Until Maturity", ascending=False)
    firm_detail["Maturity Date"] = firm_detail["Maturity Date"].dt.strftime("%d %b %Y")
    firm_detail["Invoice Value"] = firm_detail["Invoice Value"].map(lambda x: f"${x:,.2f}")
    firm_detail["Days Until Maturity"] = firm_detail["Days Until Maturity"].map(lambda x: f"{int(x)}d")
    st.dataframe(firm_detail, use_container_width=True, hide_index=True,
                 height=min(500, 38*len(firm_detail)+38))

# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — DAILY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t_daily:
    min_d = df['_date'].dt.date.min()
    max_d = df['_date'].dt.date.max()
    today_d = pd.Timestamp.today().date()
    default_sel = today_d if (min_d is not pd.NaT and max_d is not pd.NaT and min_d <= today_d <= max_d) else (max_d if pd.notna(max_d) else today_d)

    daily_sel = st.date_input(
        "📅 Select Date for Daily Analysis",
        value=default_sel,
        min_value=min_d if pd.notna(min_d) else None,
        max_value=max_d if pd.notna(max_d) else None,
        key="daily_sel_date",
    )

    df_daily = df[df["_date"].dt.normalize().dt.date == daily_sel].copy()
    daily_N = len(df_daily)

    st.caption(f"Daily data for: {pd.to_datetime(daily_sel).strftime('%d %b %Y')}  |  Records: {daily_N:,}")

    if df_daily.empty:
        st.warning("No records for the selected date (after applying global filters).")
        st.stop()

    daily_qty   = df_daily["Invoice Qty"].sum() if "Invoice Qty" in df_daily.columns else 0
    daily_avg   = df_daily["Invoice Value"].sum() / daily_N if daily_N else 0
    daily_inv   = df_daily["Invoice Value"].sum() if "Invoice Value" in df_daily.columns else 0
    paid_amt    = df_daily[df_daily["Payment. Rcv Dt"].notna()]["Invoice Value"].sum()
    pending_amt = df_daily[df_daily["Payment. Rcv Dt"].isna()]["Invoice Value"].sum()

    daily_kpi_css = """
    <style>
    .daily-3d-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 14px;
        margin: 14px 0;
    }
    .daily-3d-card {
        background: linear-gradient(145deg, #112240, #0a1628);
        border: 1px solid rgba(0,201,167,0.18);
        border-radius: 14px;
        padding: 18px 16px;
        transform: perspective(700px) rotateX(3deg) rotateY(0deg) translateZ(0);
        box-shadow: 0 8px 25px rgba(0,0,0,0.45), 0 4px 10px rgba(0,201,167,0.06),
                    inset 0 1px 0 rgba(255,255,255,0.06), inset 0 -2px 0 rgba(0,0,0,0.3);
        transition: all 0.25s cubic-bezier(.4,0,.2,1);
        position: relative;
        overflow: hidden;
    }
    .daily-3d-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
        background: linear-gradient(90deg, transparent, var(--card-accent, rgba(0,201,167,0.5)), transparent);
    }
    .daily-3d-card::after {
        content: '';
        position: absolute;
        bottom: 0; left: 10%; right: 10%;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,201,167,0.2), transparent);
    }
    .daily-3d-card:hover {
        transform: perspective(700px) rotateX(0deg) rotateY(0deg) translateZ(8px) translateY(-4px);
        box-shadow: 0 16px 40px rgba(0,201,167,0.15), 0 8px 20px rgba(0,0,0,0.5),
                    inset 0 1px 0 rgba(255,255,255,0.10);
        border-color: rgba(0,201,167,0.45);
    }
    .daily-3d-card:active {
        transform: perspective(700px) rotateX(2deg) translateZ(2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.4), inset 0 2px 4px rgba(0,0,0,0.2);
    }
    .daily-3d-icon { font-size: 18px; margin-bottom: 6px; }
    .daily-3d-label { font-size: 11px; color: #8899aa; font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; margin-bottom: 6px; }
    .daily-3d-value { font-size: 26px; font-weight: 800; color: #00c9a7; line-height: 1.1; text-shadow: 0 1px 3px rgba(0,0,0,0.4); }
    .daily-3d-delta { font-size: 11px; color: #8899aa; margin-top: 6px; font-weight: 500; }
    </style>
    """
    st.markdown(daily_kpi_css, unsafe_allow_html=True)

    k1_val = f"{daily_N:,}"
    k2_val = usd(daily_inv)
    k3_val = f"{daily_qty:,.0f}"
    k4_val = usd(daily_avg)
    k5_val = f"{df_daily['Party Name'].nunique():,}"
    k6_val = f"{df_daily['Our Bank'].nunique():,}"

    st.markdown(f"""
    <div class="daily-3d-grid">
        <div class="daily-3d-card">
            <div class="daily-3d-icon">📋</div>
            <div class="daily-3d-label">Total Submissions</div>
            <div class="daily-3d-value">{k1_val}</div>
            <div class="daily-3d-delta">{daily_N} records</div>
        </div>
        <div class="daily-3d-card">
            <div class="daily-3d-icon">💵</div>
            <div class="daily-3d-label">Invoice Value</div>
            <div class="daily-3d-value">{k2_val}</div>
            <div class="daily-3d-delta">total invoiced</div>
        </div>
        <div class="daily-3d-card">
            <div class="daily-3d-icon">📦</div>
            <div class="daily-3d-label">Invoice Qty</div>
            <div class="daily-3d-value">{k3_val}</div>
            <div class="daily-3d-delta">total quantity</div>
        </div>
        <div class="daily-3d-card">
            <div class="daily-3d-icon">📈</div>
            <div class="daily-3d-label">Avg Value</div>
            <div class="daily-3d-value">{k4_val}</div>
            <div class="daily-3d-delta">per submission</div>
        </div>
        <div class="daily-3d-card">
            <div class="daily-3d-icon">🏢</div>
            <div class="daily-3d-label">Unique Parties</div>
            <div class="daily-3d-value">{k5_val}</div>
            <div class="daily-3d-delta">active parties</div>
        </div>
        <div class="daily-3d-card">
            <div class="daily-3d-icon">🏦</div>
            <div class="daily-3d-label">Unique Banks</div>
            <div class="daily-3d-value">{k6_val}</div>
            <div class="daily-3d-delta">banks used</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")
    paid_val = usd(paid_amt)
    pend_val = usd(pending_amt)
    st.markdown(f"""
    <div class="daily-3d-grid" style="grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));">
        <div class="daily-3d-card">
            <div class="daily-3d-icon">✅</div>
            <div class="daily-3d-label">Payment Received Value</div>
            <div class="daily-3d-value">{paid_val}</div>
        </div>
        <div class="daily-3d-card">
            <div class="daily-3d-icon">⏳</div>
            <div class="daily-3d-label">Pending Invoice Value</div>
            <div class="daily-3d-value">{pend_val}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    daily_grp = (df.groupby(df["_date"].dt.date)
                   .agg(Submissions=("LC No","count"),
                        Qty=("Invoice Qty","sum") if "Invoice Qty" in df.columns else ("LC No","count"),
                        Value=("Invoice Value","sum"))
                   .reset_index().rename(columns={"_date":"Date"}))
    if not daily_grp.empty:
        daily_grp["Date"] = pd.to_datetime(daily_grp["Date"])
        daily_grp = daily_grp.sort_values("Date")
        daily_grp["Cumulative Value"] = daily_grp["Value"].cumsum()

        dc1, dc2, dc3 = st.columns([2, 2, 3])
        with dc1:
            sh("📅 Daily Summary Table")
            tbl = daily_grp.copy()
            tbl["Date"]             = tbl["Date"].dt.strftime("%d %b %Y")
            tbl["Value"]            = tbl["Value"].map(lambda x: f"${x:,.2f}")
            tbl["Cumulative Value"] = tbl["Cumulative Value"].map(lambda x: f"${x:,.2f}")
            st.dataframe(tbl, use_container_width=True, hide_index=True, height=360)

        with dc2:
            sh("📈 Daily Invoice Value")
            fig_d = go.Figure()
            fig_d.add_bar(x=daily_grp["Date"], y=daily_grp["Value"], marker_color="#1a8fff")
            fig_d.update_layout(**PL_GENERAL,
                xaxis=dict(title="Date", tickangle=-40, tickfont=dict(size=9), gridcolor="#1a2a3a"),
                yaxis=dict(title="Invoice Value (USD)", gridcolor="#1a2a3a"),
                height=320, showlegend=False)
            st.plotly_chart(fig_d, use_container_width=True)

        with dc3:
            sh("📈 Cumulative Invoice Value")
            fig_cum = go.Figure()
            fig_cum.add_scatter(x=daily_grp["Date"], y=daily_grp["Cumulative Value"],
                                mode="lines+markers",
                                line=dict(color="#00c9a7", width=2.5),
                                fill="tozeroy", fillcolor="rgba(0,201,167,0.08)")
            fig_cum.update_layout(**PL_GENERAL,
                xaxis=dict(title="Date", tickangle=-40, tickfont=dict(size=9), gridcolor="#1a2a3a"),
                yaxis=dict(title="Cumulative Value (USD)", gridcolor="#1a2a3a"),
                height=320, showlegend=False)
            st.plotly_chart(fig_cum, use_container_width=True)
    else:
        st.warning("No daily data available for the current filters.")

    st.markdown("---")

    banks_order  = ["SEBPLC","PBL","CBP","DBBL","ONE"]
    by_bank_full = (df_daily.groupby("Our Bank")
                      .agg(Value=("Invoice Value","sum"), Submissions=("LC No","count"))
                      .reset_index().sort_values("Value", ascending=False))
    parts = [by_bank_full[by_bank_full["Our Bank"] == b]
             for b in banks_order if b in by_bank_full["Our Bank"].values]
    rest  = by_bank_full[~by_bank_full["Our Bank"].isin(banks_order)]
    if not rest.empty: parts.append(rest)
    by_bank_ord = pd.concat(parts, ignore_index=True) if parts else by_bank_full

    br1, br2 = st.columns(2)
    with br1:
        sh("🏦 Our Bank Breakdown")
        fig_pie = px.pie(by_bank_ord, names="Our Bank", values="Value",
                         hole=0.48, color_discrete_sequence=C)
        fig_pie.update_layout(**PL)
        fig_pie.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig_pie, use_container_width=True)
    with br2:
        sh("🏦 Our Bank Table")
        tb = by_bank_ord.copy()
        tb["Value"] = tb["Value"].map(lambda x: f"${x:,.2f}")
        tb.columns  = ["Our Bank","Invoice Value","Submissions"]
        st.dataframe(tb, use_container_width=True, hide_index=True, height=320)

    st.markdown("---")

    by_firm_full = (df_daily.groupby("Firm Name")
                      .agg(Value=("Invoice Value","sum"), Submissions=("LC No","count"))
                      .reset_index().sort_values("Value", ascending=False))
    top_firms = by_firm_full.head(10)
    f1, f2 = st.columns(2)
    with f1:
        sh("🏢 Firm Name Breakdown")
        fig_f = px.bar(top_firms, y="Firm Name", x="Value", orientation="h",
                       color="Firm Name", color_discrete_sequence=C,
                       text=top_firms["Value"].map(usd))
        fig_f.update_traces(textposition="outside", textfont_size=10)
        fig_f.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=360)
        st.plotly_chart(fig_f, use_container_width=True)
    with f2:
        sh("🏢 Firm Name Table")
        tf = top_firms.copy()
        tf["Value"] = tf["Value"].map(lambda x: f"${x:,.2f}")
        tf.columns  = ["Firm Name","Invoice Value","Submissions"]
        st.dataframe(tf, use_container_width=True, hide_index=True, height=360)

    st.markdown("---")

    spg_d = (df_daily[df_daily["Sales Person"].notna()]
               .groupby("Sales Person")
               .agg(Value=("Invoice Value","sum"), N=("LC No","count"))
               .reset_index().sort_values("Value", ascending=False))
    if not spg_d.empty:
        sp_paid = (df_daily[df_daily["Payment. Rcv Dt"].notna() & df_daily["Sales Person"].notna()]
                     .groupby("Sales Person").size().reset_index(name="Paid"))
        spg_d = spg_d.merge(sp_paid, on="Sales Person", how="left").fillna(0)
        spg_d["Pct"] = (spg_d["Paid"] / spg_d["N"] * 100).round(1)
        total_val_d  = spg_d["Value"].sum()
        spg_d["% of Total"] = (spg_d["Value"] / total_val_d * 100).round(1).map(lambda x: f"{x:.1f}%")

        s1, s2 = st.columns(2)
        with s1:
            sh("👤 Sales Person Performance")
            sp_show = spg_d[["Sales Person","Value","N","Paid","Pct","% of Total"]].copy()
            sp_show["Value"] = sp_show["Value"].map(lambda x: f"${x:,.2f}")
            sp_show["Pct"]   = sp_show["Pct"].map(lambda x: f"{x:.1f}%")
            sp_show["Paid"]  = sp_show["Paid"].astype(int)
            sp_show.columns  = ["Sales Person","Invoice Value","Submissions","Paid","Pay Rate","% of Total"]
            st.dataframe(sp_show, use_container_width=True, hide_index=True, height=360)
        with s2:
            sh("👤 Sales Person Chart")
            fig_sp = px.bar(spg_d.head(12), x="Value", y="Sales Person", orientation="h",
                            color="Sales Person", color_discrete_sequence=C,
                            text=spg_d.head(12)["Value"].map(usd))
            fig_sp.update_traces(textposition="outside", textfont_size=10)
            fig_sp.update_layout(**PL_GENERAL,
                xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
                yaxis=dict(title="", autorange="reversed"),
                showlegend=False, height=360)
            st.plotly_chart(fig_sp, use_container_width=True)

    st.markdown("---")

    ten_dist = pd.DataFrame()
    if "Tenor" in df_daily.columns:
        ten = df_daily["Tenor"].apply(norm_tenor)
        ten_dist = ten.value_counts().reset_index()
        ten_dist.columns = ["Tenor","Count"]

    t1a, t1b = st.columns(2)
    with t1a:
        sh("⏱ Tenor Distribution")
        if not ten_dist.empty:
            fig_t = px.pie(ten_dist, names="Tenor", values="Count",
                           color_discrete_sequence=C, hole=0.45)
            fig_t.update_layout(**PL)
            fig_t.update_traces(textinfo="label+percent", textfont_size=11)
            st.plotly_chart(fig_t, use_container_width=True)
        else:
            st.warning("Tenor column not available.")
    with t1b:
        sh("⏱ Tenor Distribution Table")
        if not ten_dist.empty:
            st.dataframe(ten_dist, use_container_width=True, hide_index=True, height=300)

    st.markdown("---")

    t_party_d = (df_daily.groupby("Party Name")
                   .agg(Value=("Invoice Value","sum"), N=("LC No","count"))
                   .reset_index().sort_values("Value", ascending=False).head(10))
    p1, p2 = st.columns(2)
    with p1:
        sh("🏭 Top 10 Party Names")
        fig_tp = px.bar(t_party_d.sort_values("Value"), x="Value", y="Party Name",
                        orientation="h", color="Party Name", color_discrete_sequence=C,
                        text=t_party_d["Value"].map(usd))
        fig_tp.update_traces(textposition="outside", textfont_size=9)
        fig_tp.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False, height=360)
        st.plotly_chart(fig_tp, use_container_width=True)
    with p2:
        sh("🏭 Top 10 Party Names Table")
        tbp = t_party_d.copy()
        tbp.insert(0, "Rank", range(1, len(tbp)+1))
        tbp["Value"] = tbp["Value"].map(lambda x: f"${x:,.2f}")
        tbp.columns  = ["Rank","Party Name","Invoice Value","Submissions"]
        st.dataframe(tbp, use_container_width=True, hide_index=True, height=360)

    st.markdown("---")

    buyer_bank = (df_daily.groupby("Bank Name")
                    .agg(Value=("Invoice Value","sum"), Submissions=("LC No","count"))
                    .reset_index().sort_values("Value", ascending=False).head(23))
    b1, b2 = st.columns(2)
    with b1:
        sh("🏛 Buyer's Bank Breakdown")
        fig_bb = px.bar(buyer_bank, x="Value", y="Bank Name", orientation="h",
                        color_discrete_sequence=[C[4]],
                        text=buyer_bank["Value"].map(usd))
        fig_bb.update_traces(textposition="outside", textfont_size=8, marker_color=C[4])
        fig_bb.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=500)
        st.plotly_chart(fig_bb, use_container_width=True)
    with b2:
        sh("🏛 Buyer's Bank Table")
        tbob = buyer_bank.copy()
        total_bb = tbob["Value"].sum()
        tbob["% Share"] = (tbob["Value"] / total_bb * 100).map(lambda x: f"{x:.1f}%")
        tbob["Value"]   = tbob["Value"].map(lambda x: f"${x:,.2f}")
        tbob.columns    = ["Bank Name","Invoice Value","Submissions","% Share"]
        st.dataframe(tbob, use_container_width=True, hide_index=True, height=500)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW (with Forecasting)
# ══════════════════════════════════════════════════════════════════════════════
with t_overview:
    l, r = st.columns(2)
    with l:
        sh("📅 Monthly Submission Trend")
        fig = go.Figure()
        fig.add_bar(x=monthly["Month"], y=monthly["Count"], name="Submissions",
                    marker_color="#1a8fff", yaxis="y1")
        fig.add_scatter(x=monthly["Month"], y=monthly["Inv"], name="Invoice Value",
                        mode="lines+markers", line=dict(color="#00c9a7", width=2.5),
                        marker=dict(size=6), yaxis="y2")
        fig.update_layout(**PL_GENERAL,
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            yaxis2=dict(title="Invoice Value (USD)", overlaying="y", side="right",
                        gridcolor="rgba(0,0,0,0)", tickformat="$.2s"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#8899aa", size=10)),
            height=340)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("🏦 Our Bank — Invoice Value")
        fig2 = px.bar(by_bank, x="Our Bank", y="Inv", color="Our Bank",
                      color_discrete_sequence=C, text=by_bank["Inv"].apply(usd))
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(**PL, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    a, b_col = st.columns(2)
    with a:
        sh("🏢 Firm-wise Invoice Value")
        fig3 = px.pie(by_firm, names="Firm Name", values="Inv",
                      color_discrete_sequence=C, hole=0.45)
        fig3.update_layout(**PL)
        fig3.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig3, use_container_width=True)
    with b_col:
        sh("📋 Monthly Summary Table")
        tbl = monthly[["MonthSort","Month","Count","Inv","LC"]].copy()
        paid_value = df[df["Payment. Rcv Dt"].notna()].groupby("MonthSort")["Invoice Value"].sum()
        pending_value = df[df["Payment. Rcv Dt"].isna()].groupby("MonthSort")["Invoice Value"].sum()
        accepted_value = df[(df["Payment. Rcv Dt"].isna()) & (df["Bank Accept Date"].notna())].groupby("MonthSort")["Invoice Value"].sum()
        not_accepted_value = df[df["Bank Accept Date"].isna()].groupby("MonthSort")["Invoice Value"].sum()
        paid_cnt = df[df["Payment. Rcv Dt"].notna()].groupby("MonthSort").size()
        accepted_cnt = df[(df["Payment. Rcv Dt"].isna()) & (df["Bank Accept Date"].notna())].groupby("MonthSort").size()
        not_accepted_cnt = df[df["Bank Accept Date"].isna()].groupby("MonthSort").size()

        tbl = (tbl.merge(paid_value.rename("Paid Value").reset_index(), on="MonthSort", how="left")
                  .merge(pending_value.rename("Pending Value").reset_index(), on="MonthSort", how="left")
                  .merge(accepted_value.rename("Accepted Value").reset_index(), on="MonthSort", how="left")
                  .merge(not_accepted_value.rename("Not Accepted Value").reset_index(), on="MonthSort", how="left")
                  .merge(paid_cnt.rename("Paid Count").reset_index(), on="MonthSort", how="left")
                  .merge(accepted_cnt.rename("Accepted Count").reset_index(), on="MonthSort", how="left")
                  .merge(not_accepted_cnt.rename("Not Accepted Count").reset_index(), on="MonthSort", how="left"))

        for c in ["Paid Value","Pending Value","Accepted Value","Not Accepted Value","Paid Count","Accepted Count","Not Accepted Count"]:
            tbl[c] = tbl[c].fillna(0)

        tbl["Paid Count %"] = (tbl["Paid Count"] / tbl["Count"] * 100).replace([pd.NA, pd.NaT, float("inf")], 0).round(1)
        tbl = tbl[["Month","Count","Inv","LC","Paid Count","Paid Count %","Paid Value","Accepted Count","Accepted Value","Not Accepted Value"]].copy()
        tbl.columns = ["Month","Submissions","Invoice Value (USD)","LC Value (USD)","Paid Count","Paid Count %","Paid Value (USD)","Accepted Count","Accepted Value (USD)","Not Accepted Value (USD)"]
        for c in ["Invoice Value (USD)","LC Value (USD)","Paid Value (USD)","Accepted Value (USD)","Not Accepted Value (USD)"]:
            tbl[c] = tbl[c].map(lambda x: f"${x:,.2f}")
        tbl["Paid Count"] = tbl["Paid Count"].astype(int)
        tbl["Accepted Count"] = tbl["Accepted Count"].astype(int)
        tbl["Paid Count %"] = tbl["Paid Count %"].map(lambda x: f"{x:.1f}%")
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # ── FORECASTING SECTION ───────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📈 Next Month Forecast (Linear Trend)")

    if len(monthly) >= 3:
        x_vals = np.arange(len(monthly))
        y_count = monthly["Count"].values.astype(float)
        y_inv   = monthly["Inv"].values.astype(float)

        # Linear regression
        def linreg(x, y):
            n = len(x)
            if n < 2: return 0, y[-1] if len(y) else 0
            sx, sy = x.sum(), y.sum()
            sxx = (x*x).sum()
            sxy = (x*y).sum()
            denom = n*sxx - sx*sx
            if denom == 0: return 0, sy/n
            m = (n*sxy - sx*sy) / denom
            b = (sy - m*sx) / n
            return m, b

        m_c, b_c = linreg(x_vals, y_count)
        m_i, b_i = linreg(x_vals, y_inv)

        next_x = len(monthly)
        pred_count = max(0, m_c * next_x + b_c)
        pred_inv   = max(0, m_i * next_x + b_i)

        # Confidence (residual std)
        pred_hist_c = m_c * x_vals + b_c
        pred_hist_i = m_i * x_vals + b_i
        std_c = np.std(y_count - pred_hist_c)
        std_i = np.std(y_inv - pred_hist_i)

        last_month_period = monthly["MonthSort"].iloc[-1]
        try:
            next_month_str = (last_month_period + 1).strftime("%b %Y")
        except Exception:
            next_month_str = "Next Month"

        fc1, fc2, fc3, fc4 = st.columns(4)
        fc1.metric("📅 Forecast Month", next_month_str)
        fc2.metric("📋 Predicted Submissions", f"{pred_count:.0f}",
                   delta=f"±{std_c:.0f} (1σ)")
        fc3.metric("💵 Predicted Invoice Value", usd(pred_inv),
                   delta=f"±{usd(std_i)} (1σ)")
        trend_dir = "↑ Upward" if m_i > 0 else "↓ Downward"
        fc4.metric("📊 Trend Direction", trend_dir)

        # Forecast chart
        all_months = list(monthly["Month"]) + [next_month_str]
        all_hist_c = list(y_count) + [None]
        all_pred_c = list(pred_hist_c) + [pred_count]

        fig_fc = go.Figure()
        fig_fc.add_scatter(x=monthly["Month"], y=y_count, name="Actual Submissions",
                           mode="lines+markers", line=dict(color="#1a8fff", width=2.5),
                           marker=dict(size=6))
        fig_fc.add_scatter(x=all_months, y=all_pred_c, name="Trend Line",
                           mode="lines+markers",
                           line=dict(color="#ff6b35", width=2, dash="dot"),
                           marker=dict(size=7, symbol="diamond",
                                       color=["rgba(0,0,0,0)"]*len(monthly) + ["#ff6b35"]))
        # Confidence band for next month
        fig_fc.add_scatter(
            x=[next_month_str, next_month_str],
            y=[max(0, pred_count - std_c), pred_count + std_c],
            mode="lines", line=dict(color="#ff6b35", width=0),
            fill="tonexty", fillcolor="rgba(255,107,53,0.15)",
            name="Confidence", showlegend=True
        )
        fig_fc.update_layout(**PL_GENERAL,
            xaxis=dict(title="Month", gridcolor="#1a2a3a"),
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=10)),
            height=320)
        st.plotly_chart(fig_fc, use_container_width=True)

        st.caption("📌 Forecast based on linear trend from historical data. Actual results may vary.")
    else:
        st.info("⚡ Need at least 3 months of data for forecasting.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — WEEKLY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t_weekly:
    best_wk  = weekly.loc[weekly["Count"].idxmax()]
    best_inv = weekly.loc[weekly["Inv"].idxmax()]

    w1, w2, w3, w4 = st.columns(4)
    w1.metric("📆 Total Weeks",       f"{len(weekly)}")
    w2.metric("🔝 Best Week (Subs)",  f"{best_wk['Week']} — {int(best_wk['Count'])}")
    w3.metric("💰 Best Week (Value)", f"{best_inv['Week']} — {usd(best_inv['Inv'])}")
    w4.metric("📊 Weekly Avg Subs",   f"{weekly['Count'].mean():.0f}")
    st.markdown("")

    sh("📅 Week-wise Submission Count + Invoice Value")
    fig = go.Figure()
    fig.add_bar(x=weekly["Week"], y=weekly["Count"], name="Submissions",
                marker_color="#1a8fff", yaxis="y1",
                text=weekly["Count"], textposition="outside", textfont=dict(size=9))
    fig.add_scatter(x=weekly["Week"], y=weekly["Inv"], name="Invoice Value (USD)",
                    mode="lines+markers", line=dict(color="#00c9a7", width=2.5),
                    marker=dict(size=5), yaxis="y2")
    fig.update_layout(**PL_GENERAL,
        yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
        yaxis2=dict(title="Invoice Value (USD)", overlaying="y", side="right",
                    gridcolor="rgba(0,0,0,0)", tickformat="$.2s"),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#8899aa", size=10)),
        height=340)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("---")

    l, r = st.columns(2)
    with l:
        sh("🏢 Weekly Invoice Value by Firm (Stacked)")
        fig2 = px.bar(wk_firm, x="Week", y="Inv", color="Firm Name",
                      color_discrete_sequence=C, barmode="stack")
        fig2.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig2, use_container_width=True)
    with r:
        sh("🏦 Weekly Invoice Value by Our Bank (Stacked)")
        fig3 = px.bar(wk_bank, x="Week", y="Inv", color="Our Bank",
                      color_discrete_sequence=C, barmode="stack")
        fig3.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig3, use_container_width=True)
    st.markdown("---")

    l2, r2 = st.columns(2)
    with l2:
        sh("🧑‍💼 Weekly Invoice by Sales Person (Top 6)")
        top6_sp = spg.head(6)["Sales Person"].tolist()
        wk_sp6  = wk_sp[wk_sp["Sales Person"].isin(top6_sp)]
        fig4 = px.line(wk_sp6, x="Week", y="Inv", color="Sales Person",
                       color_discrete_sequence=C, markers=True)
        fig4.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.15, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig4, use_container_width=True)
    with r2:
        sh("🔄 Weekly Payment Status (Stacked)")
        fig5 = px.bar(wk_st_grp, x="Week", y="Count", color="Status",
                      barmode="stack",
                      color_discrete_map={"Paid":"#00c9a7","Accepted":"#1a8fff","Not Accepted":"#ff6b35"})
        fig5.update_layout(**PL_GENERAL,
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=320)
        st.plotly_chart(fig5, use_container_width=True)
    st.markdown("---")

    sh("👥 Weekly Invoice Value — Top 5 Parties (Line)")
    wk_p5 = wk_party_top[wk_party_top["Party Name"].isin(t_party.head(5)["Party Name"].tolist())]
    fig6  = px.line(wk_p5, x="Week", y="Inv", color="Party Name",
                    color_discrete_sequence=C, markers=True)
    fig6.update_layout(**PL_GENERAL,
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), title="", gridcolor="#1a2a3a"),
        yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
        height=300)
    st.plotly_chart(fig6, use_container_width=True)
    st.markdown("---")

    sh("📋 Weekly Summary Table")
    wk_status2 = df[["Week","Payment. Rcv Dt","Bank Accept Date","Invoice Value"]].copy()
    wk_status2["Status"] = wk_status2.apply(
        lambda r: "Paid" if pd.notna(r["Payment. Rcv Dt"]) else
                  ("Accepted" if pd.notna(r["Bank Accept Date"]) else "Not Accepted"), axis=1)
    total2 = wk_status2.groupby("Week").size().reset_index(name="Submissions")
    paid_cnt2 = wk_status2[wk_status2["Status"] == "Paid"].groupby("Week").size().reset_index(name="Paid")
    val2 = wk_status2.groupby(["Week","Status"])["Invoice Value"].sum().reset_index()
    val2 = val2.pivot(index="Week", columns="Status", values="Invoice Value").reset_index()
    val2 = val2.rename(columns={"Paid":"Paid Value","Accepted":"Accepted Value","Not Accepted":"Not Accepted Value"})
    for _col in ["Paid Value","Accepted Value","Not Accepted Value"]:
        if _col not in val2.columns:
            val2[_col] = 0
    summary2 = total2.merge(paid_cnt2, on="Week", how="left").merge(val2, on="Week", how="left").fillna(0)
    summary2["Payment Rate"] = (summary2["Paid"] / summary2["Submissions"] * 100).round(1)
    total_val2 = wk_status2.groupby("Week")["Invoice Value"].sum().reset_index(name="Invoice Value (USD)")
    summary2 = summary2.merge(total_val2, on="Week", how="left")
    summary2 = summary2[["Week","Submissions","Invoice Value (USD)","Paid","Payment Rate","Paid Value","Accepted Value","Not Accepted Value"]]

    subs_total2 = int(summary2["Submissions"].sum())
    paid_total2 = int(summary2["Paid"].sum())
    inv_total2  = float(summary2["Invoice Value (USD)"].sum())
    paid_val_total2 = float(summary2["Paid Value"].sum())
    acc_val_total2  = float(summary2["Accepted Value"].sum())
    nacc_val_total2 = float(summary2["Not Accepted Value"].sum())
    pay_rate_total2 = f"{(paid_total2 / subs_total2 * 100 if subs_total2 else 0):.1f}%"

    totals_row2 = {"Week":"TOTAL","Submissions":subs_total2,"Invoice Value (USD)":inv_total2,
                   "Paid":paid_total2,"Payment Rate":float(pay_rate_total2.replace("%","")),
                   "Paid Value":paid_val_total2,"Accepted Value":acc_val_total2,"Not Accepted Value":nacc_val_total2}
    summary_total2 = pd.concat([summary2, pd.DataFrame([totals_row2])], ignore_index=True)
    for c in ["Invoice Value (USD)","Paid Value","Accepted Value","Not Accepted Value"]:
        summary_total2[c] = summary_total2[c].map(lambda x: f"${x:,.2f}")
    summary_total2["Payment Rate"] = summary_total2["Payment Rate"].map(lambda x: f"{float(x):.1f}%")
    summary_total2["Paid"] = summary_total2["Paid"].astype(int)
    st.dataframe(summary_total2, use_container_width=True, hide_index=True)

    if REPORTLAB_AVAILABLE:
        def weekly_summary_to_pdf_bytes(df_in, title="Bank Submissions Weekly Summary Table 2026", subtitle="", generated_at=""):
            buf = BytesIO()
            pw, ph = _landscape(A3)
            lm = rm = 24; tm = bm = 24; uw = pw - lm - rm
            hf = "Helvetica-Bold"; cf = "Helvetica"; hfs = 12; cfs = 9
            min_cw = 1.0 * inch; max_cw = 3.5 * inch

            def mw(txt, font, sz): return pdfmetrics.stringWidth(str(txt), font, sz)

            dt = df_in.fillna("").astype(str)
            col_widths_pdf = []
            for col in df_in.columns:
                vals = dt[col].tolist()
                if len(vals) > 200: vals = vals[::max(1, len(vals)//200)]
                measured = [mw(v, cf, cfs) for v in vals if v]
                mx = max([mw(col, hf, hfs)] + measured) if measured else mw(col, hf, hfs)
                col_widths_pdf.append(max(min_cw, min(max_cw, mx + 18)))

            tot = sum(col_widths_pdf)
            if tot > uw and tot > 0:
                col_widths_pdf = [w * uw / tot for w in col_widths_pdf]

            def chunk(cols, widths, max_w):
                groups, cur, cw3 = [], [], 0.0
                for c, w in zip(cols, widths):
                    if cur and cw3 + w > max_w: groups.append(cur); cur = [c]; cw3 = w
                    else: cur.append(c); cw3 += w
                if cur: groups.append(cur)
                return groups

            ss2 = getSampleStyleSheet()
            hs = ParagraphStyle("HS", parent=ss2["Normal"], fontName=hf, fontSize=hfs, leading=13,
                                textColor=colors.HexColor("#ffffff"), alignment=TA_LEFT, wordWrap="CJK")
            cs = ParagraphStyle("CS", parent=ss2["Normal"], fontName=cf, fontSize=cfs, leading=10,
                                alignment=TA_LEFT, wordWrap="CJK")
            ts = TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0d3f47")),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("ALIGN",(0,0),(-1,-1),"LEFT"), ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),8),
                ("LEFTPADDING",(0,0),(-1,-1),4), ("RIGHTPADDING",(0,0),(-1,-1),4),
                ("BOTTOMPADDING",(0,0),(-1,-1),3), ("TOPPADDING",(0,0),(-1,-1),3),
                ("GRID",(0,0),(-1,-1),0.25,colors.grey), ("BOX",(0,0),(-1,-1),0.5,colors.black),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.whitesmoke,colors.lightgrey]),
            ])
            col_list = list(df_in.columns)
            groups = chunk(col_list, col_widths_pdf, uw)
            pages = []
            for g_i, g_cols in enumerate(groups):
                g_w = [col_widths_pdf[col_list.index(c)] for c in g_cols]
                rows = [[Paragraph(str(c), hs) for c in g_cols]]
                for rv in df_in.fillna("").astype(str).values.tolist():
                    rows.append([Paragraph(str(rv[col_list.index(c)]), cs) for c in g_cols])
                tbl2 = LongTable(rows, repeatRows=1, colWidths=g_w, hAlign="CENTER", splitByRow=1, spaceBefore=12, spaceAfter=12)
                tbl2.setStyle(ts)
                pages.append(tbl2)
                if g_i < len(groups) - 1: pages.append(PageBreak())

            def page_header(canvas, doc):
                canvas.saveState()
                canvas.setFont(hf, 18); canvas.drawCentredString(pw/2, ph-tm-8, str(title))
                canvas.setFont(cf, 11)
                if subtitle: canvas.drawCentredString(pw/2, ph-tm-26, str(subtitle))
                if generated_at: canvas.setFont(cf, 8); canvas.drawString(lm, bm/2+2, f"Generated on: {generated_at}")
                canvas.setFont(cf, 8); canvas.drawRightString(pw-rm, ph-tm-8, f"Page {doc.page}")
                canvas.restoreState()

            doc = SimpleDocTemplate(buf, pagesize=_landscape(A3), leftMargin=lm, rightMargin=rm, topMargin=tm+28, bottomMargin=bm)
            doc.build(pages, onFirstPage=page_header, onLaterPages=page_header)
            buf.seek(0); return buf.read()

        export_weekly = summary_total2.copy()[["Week","Submissions","Invoice Value (USD)","Paid","Payment Rate","Paid Value","Accepted Value","Not Accepted Value"]]
        pdf_bytes_weekly = weekly_summary_to_pdf_bytes(
            export_weekly,
            subtitle=f"Generated from current filters ({len(summary2)} weeks + TOTAL row)",
            generated_at=datetime.now().strftime("%d %b %Y %H:%M:%S"))
        st.download_button("📄 Download Weekly Summary PDF", pdf_bytes_weekly,
            file_name="bank_submissions_weekly_summary.pdf", mime="application/pdf")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — FIRM & SALES PERSON
# ══════════════════════════════════════════════════════════════════════════════
with t_firm:
    l, r = st.columns(2)
    with l:
        sh("🏢 Firm-wise Invoice Value")
        fig = px.bar(by_firm, y="Firm Name", x="Inv", orientation="h",
                     color="Firm Name", color_discrete_sequence=C,
                     text=by_firm["Inv"].apply(usd))
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

        sh("Firm Ranking Table")
        ft = by_firm.copy()
        ft["% Share"] = (ft["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        ft["Inv"]     = ft["Inv"].map(lambda x: f"${x:,.2f}")
        ft.columns    = ["Firm","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(ft, use_container_width=True, hide_index=True)
    with r:
        sh("🧑‍💼 Sales Person — Invoice Value (Top 12)")
        fig2 = px.bar(spg.head(12), y="Sales Person", x="Inv", orientation="h",
                      color="Sales Person", color_discrete_sequence=C,
                      text=spg.head(12)["Inv"].apply(usd))
        fig2.update_traces(textposition="outside", textfont_size=10)
        fig2.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title=""), showlegend=False, height=400)
        st.plotly_chart(fig2, use_container_width=True)

        sh("Sales Person Ranking Table")
        sp_rank = spg[["Sales Person","Inv","N","Paid","Pct"]].copy()
        sp_accepted = (df[df["Payment. Rcv Dt"].isna() & df["Bank Accept Date"].notna()]
                         .groupby("Sales Person")["Invoice Value"].sum())
        sp_not_accepted = (df[df["Bank Accept Date"].isna()]
                            .groupby("Sales Person")["Invoice Value"].sum())
        sp_rank["Accepted Value"] = sp_rank["Sales Person"].map(sp_accepted).fillna(0.0)
        sp_rank["Not Accepted Value"] = sp_rank["Sales Person"].map(sp_not_accepted).fillna(0.0)
        sp_paid_v = (df[df["Payment. Rcv Dt"].notna()].groupby("Sales Person")["Invoice Value"].sum())
        sp_rank["Paid Value"] = sp_rank["Sales Person"].map(sp_paid_v).fillna(0.0)

        sp_rank["Inv"]              = sp_rank["Inv"].map(lambda x: f"${x:,.2f}")
        sp_rank["Paid Value"]       = sp_rank["Paid Value"].map(lambda x: f"${x:,.2f}")
        sp_rank["Accepted Value"]   = sp_rank["Accepted Value"].map(lambda x: f"${x:,.2f}")
        sp_rank["Not Accepted Value"] = sp_rank["Not Accepted Value"].map(lambda x: f"${x:,.2f}")
        sp_rank["Pct"] = sp_rank["Pct"].map(lambda x: f"{x:.1f}%")
        sp_rank["Paid"] = sp_rank["Paid"].astype(int)
        st2 = sp_rank[["Sales Person","Inv","N","Paid","Pct","Paid Value","Accepted Value","Not Accepted Value"]].copy()
        st2.columns = ["Sales Person","Invoice Value (USD)","Submissions","Paid","Payment Rate","Paid Value","Accepted Value","Not Accepted Value"]
        st.dataframe(st2, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — BANKS
# ══════════════════════════════════════════════════════════════════════════════
with t_banks:
    l, r = st.columns(2)
    with l:
        sh("🏦 Our Bank — Share (Donut)")
        fig = px.pie(by_bank, names="Our Bank", values="Inv",
                     color_discrete_sequence=C, hole=0.48)
        fig.update_layout(**PL)
        fig.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("🏛️ Top 10 Party Banks")
        fig2 = px.bar(t_bname, y="Bank Name", x="Inv", orientation="h",
                      text=t_bname["Inv"].apply(usd), color_discrete_sequence=[C[4]])
        fig2.update_traces(textposition="outside", textfont_size=10, marker_color=C[4])
        fig2.update_layout(**PL, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    a2, b2 = st.columns(2)
    with a2:
        sh("Our Bank Detail")
        ob = by_bank.copy()
        ob["% Share"] = (ob["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        ob["Inv"]     = ob["Inv"].map(lambda x: f"${x:,.2f}")
        ob.columns    = ["Our Bank","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(ob, use_container_width=True, hide_index=True)
    with b2:
        sh("Top Party Banks Detail")
        pb = t_bname.copy()
        pb["% Share"] = (pb["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        pb["Inv"]     = pb["Inv"].map(lambda x: f"${x:,.2f}")
        pb.columns    = ["Bank Name","Invoice Value (USD)","Submissions","% Share"]
        st.dataframe(pb, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — TOP PARTIES
# ══════════════════════════════════════════════════════════════════════════════
with t_parties:
    l, r = st.columns([3, 2])
    with l:
        sh("👥 Top 10 Party — Invoice Value")
        fig = px.bar(t_party, y="Party Name", x="Inv", orientation="h",
                     color="Party Name", color_discrete_sequence=C,
                     text=t_party["Inv"].apply(usd))
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            showlegend=False, height=400)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        sh("Party Ranking")
        pt = t_party.copy()
        pt.insert(0, "#", range(1, len(pt)+1))
        pt["% Share"] = (pt["Inv"]/inv*100).map(lambda x: f"{x:.1f}%")
        pt["Inv"]     = pt["Inv"].map(lambda x: f"${x:,.2f}")
        pt            = pt[["#","Party Name","Inv","N","% Share"]]
        pt.columns    = ["#","Party","Invoice Value","Subs","% Share"]
        st.dataframe(pt, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — PAYMENT STATUS
# ══════════════════════════════════════════════════════════════════════════════
with t_payment:
    pay_kpi_css = """
    <style>
    .pay-3d-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 14px;
        margin: 14px 0;
    }
    .pay-3d-card {
        background: linear-gradient(145deg, #112240, #0a1628);
        border: 1px solid rgba(0,201,167,0.18);
        border-radius: 14px;
        padding: 22px 20px;
        transform: perspective(700px) rotateX(3deg) rotateY(0deg) translateZ(0);
        box-shadow: 0 8px 25px rgba(0,0,0,0.45), 0 4px 10px rgba(0,201,167,0.06),
                    inset 0 1px 0 rgba(255,255,255,0.06), inset 0 -2px 0 rgba(0,0,0,0.3);
        transition: all 0.25s cubic-bezier(.4,0,.2,1);
        position: relative;
        overflow: hidden;
    }
    .pay-3d-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
        background: linear-gradient(90deg, transparent, var(--card-accent, rgba(0,201,167,0.5)), transparent);
    }
    .pay-3d-card::after {
        content: '';
        position: absolute;
        bottom: 0; left: 10%; right: 10%;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,201,167,0.2), transparent);
    }
    .pay-3d-card:hover {
        transform: perspective(700px) rotateX(0deg) rotateY(0deg) translateZ(8px) translateY(-4px);
        box-shadow: 0 16px 40px rgba(0,201,167,0.15), 0 8px 20px rgba(0,0,0,0.5),
                    inset 0 1px 0 rgba(255,255,255,0.10);
        border-color: rgba(0,201,167,0.45);
    }
    .pay-3d-card:active {
        transform: perspective(700px) rotateX(2deg) translateZ(2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.4), inset 0 2px 4px rgba(0,0,0,0.2);
    }
    .pay-3d-icon { font-size: 22px; margin-bottom: 6px; }
    .pay-3d-label { font-size: 12px; color: #8899aa; font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; margin-bottom: 6px; }
    .pay-3d-count { font-size: 36px; font-weight: 800; color: #e2eaf3; line-height: 1.1; text-shadow: 0 1px 3px rgba(0,0,0,0.4); }
    .pay-3d-delta { font-size: 12px; margin-top: 8px; font-weight: 600; display: flex; align-items: center; gap: 4px; }
    .pay-3d-delta.green { color: #00c9a7; }
    .pay-3d-delta.blue { color: #1a8fff; }
    .pay-3d-delta.red { color: #ff3b30; }
    </style>
    """
    st.markdown(pay_kpi_css, unsafe_allow_html=True)

    paid_pct = f"{paid_n/N*100:.1f}%" if N else "0%"
    acc_pct = f"{acc_n/N*100:.1f}%" if N else "0%"
    nacc_pct = f"{nacc_n/N*100:.1f}%" if N else "0%"

    st.markdown(f"""
    <div class="pay-3d-grid">
        <div class="pay-3d-card" style="--card-accent: rgba(0,201,167,0.6);">
            <div class="pay-3d-icon">✅</div>
            <div class="pay-3d-label">Payment Received</div>
            <div class="pay-3d-count">{paid_n:,}</div>
            <div class="pay-3d-delta green">↑ {usd(pay_v)} · {paid_pct}</div>
        </div>
        <div class="pay-3d-card" style="--card-accent: rgba(26,143,255,0.6);">
            <div class="pay-3d-icon">⏳</div>
            <div class="pay-3d-label">Accepted — Pending Pmt</div>
            <div class="pay-3d-count">{acc_n:,}</div>
            <div class="pay-3d-delta blue">↑ {usd(acc_v)} · {acc_pct}</div>
        </div>
        <div class="pay-3d-card" style="--card-accent: rgba(255,107,53,0.6);">
            <div class="pay-3d-icon">❌</div>
            <div class="pay-3d-label">Not Accepted</div>
            <div class="pay-3d-count">{nacc_n:,}</div>
            <div class="pay-3d-delta red">↑ {usd(nacc_v)} · {nacc_pct}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("")

    st_df = pd.DataFrame({
        "Status": ["Payment Received","Accepted (Pending Pmt)","Not Accepted"],
        "Count":  [paid_n, acc_n, nacc_n],
        "Value":  [pay_v,  acc_v, nacc_v],
    })
    p1, p2 = st.columns(2)
    with p1:
        sh("By Count")
        fig = px.pie(st_df, names="Status", values="Count", hole=0.5,
                     color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"])
        fig.update_layout(**PL)
        fig.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    with p2:
        sh("By Invoice Value (USD)")
        fig2 = px.pie(st_df, names="Status", values="Value", hole=0.5,
                      color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"])
        fig2.update_layout(**PL, showlegend=False)
        fig2.update_traces(textinfo="label+percent", textfont_size=11)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    sh("Sales Person — Paid vs Not Yet Paid")
    fig3 = go.Figure()
    fig3.add_bar(name="Paid",         x=spg["Sales Person"], y=spg["Paid"],          marker_color="#00c9a7")
    fig3.add_bar(name="Not Yet Paid", x=spg["Sales Person"], y=spg["N"]-spg["Paid"], marker_color="#095e59")
    fig3.update_layout(**PL_GENERAL, barmode="stack",
        yaxis=dict(title="Submissions", gridcolor="#1a2a3a"),
        xaxis=dict(title="", tickangle=-35, gridcolor="#1a2a3a"),
        legend=dict(orientation="h", x=0, y=1.1, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#8899aa", size=10)))
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    sh("Full Record Table with Search")
    search_text = st.text_input("🔍 Search all columns", key="global_search")
    pft = df.copy()
    pft["Status"] = pft.apply(lambda r:
        "✅ Paid"        if pd.notna(r["Payment. Rcv Dt"]) else
        "⏳ Accepted"   if pd.notna(r["Bank Accept Date"]) else
        "❌ Not Accepted", axis=1)
    date_cols = ["Bank Submition Date","Bank Ref Date","Lc Date","Bank Accept Date",
                 "Maturity Date","Payment. Rcv Dt","Date","Invoice Date"]
    date_cols = list(dict.fromkeys(date_cols))

    if search_text:
        pft_for_search = pft.copy()
        for col in date_cols:
            if col in pft_for_search.columns:
                pft_for_search[col] = (pd.to_datetime(pft_for_search[col], errors="coerce")
                                       .dt.strftime("%d %b %Y").fillna(""))
        tmp = pft_for_search.astype(str)
        mask = tmp.apply(lambda c: c.str.contains(search_text, case=False, na=False)).any(axis=1)
        pft = pft[mask]

    for col in date_cols:
        if col in pft.columns:
            dtcol = pd.to_datetime(pft[col], errors="coerce")
            pft[col + "__dt"] = dtcol
            pft[col] = dtcol

    internal_cols = ["_date","MonthSort","Month","WeekSort","Week","DayName"]
    hidden_dt_cols = [c + "__dt" for c in date_cols if c in pft.columns]
    internal_cols = internal_cols + hidden_dt_cols
    col_order = [
        "Firm Name","Our Bank","Bank Submition Date","Bank Ref Date","Bank Refno",
        "Party Name","LC No","Lc Date","Tenor","Bank Name","Invoice No","Invoice Date",
        "Invoice Qty","Invoice Value","Bank Accept Date","Maturity Date",
        "Payment. Rcv Dt","Sales Person","Week","DayName","Date","Status",
    ]
    export_cols = [c for c in col_order if c in pft.columns]
    extra_cols  = [c for c in pft.columns if c not in export_cols and c not in internal_cols and c not in hidden_dt_cols]
    pft_export  = pft[export_cols + extra_cols]
    pft_display = pft_export.copy()

    total_invoice_value = 0.0
    if "Invoice Value" in pft_export.columns:
        try: total_invoice_value = float(pft_export["Invoice Value"].sum())
        except Exception: pass

    def make_col_widths(dataframe, min_w=100, max_w=450, cw=8):
        text_df = dataframe.fillna("").astype(str)
        widths  = {}
        for col in text_df.columns:
            try:
                max_len = float(text_df[col].str.len().max())
                if pd.isna(max_len): max_len = 0.0
            except Exception: max_len = 0.0
            w = max(len(str(col)), int(max_len)) * cw + 24
            widths[col] = min(max_w, max(min_w, w))
        if "Bank Refno"    in widths: widths["Bank Refno"]    = max(widths["Bank Refno"],    285)
        if "LC No"         in widths: widths["LC No"]         = max(widths["LC No"],         285)
        if "Invoice Value" in widths: widths["Invoice Value"] = max(widths["Invoice Value"], 160)
        return widths

    col_widths  = make_col_widths(pft_display)
    col_config  = {}
    date_cols_set = set(date_cols)
    for c in pft_display.columns:
        if c == "Invoice Value" and pd.api.types.is_numeric_dtype(pft_display[c]):
            col_config[c] = st.column_config.NumberColumn(format="$%0.2f", width=col_widths[c])
        elif c in date_cols_set:
            col_config[c] = st.column_config.DateColumn(width=col_widths[c], format="DD MMM YYYY")
        elif pd.api.types.is_numeric_dtype(pft_display[c]):
            col_config[c] = st.column_config.NumberColumn(width=col_widths[c])
        else:
            col_config[c] = st.column_config.TextColumn(width=col_widths[c])

    st.dataframe(pft_display, use_container_width=True, hide_index=True, height=500, column_config=col_config)
    if "Invoice Value" in pft_export.columns:
        st.markdown(f"**Total Invoice Value (filtered):** ${total_invoice_value:,.2f}")

    col_csv, col_pdf = st.columns(2)
    col_csv.download_button("📥 Download CSV",
        pft_export.to_csv(index=False).encode("utf-8"),
        "bank_submit_filtered.csv", "text/csv")

    pdf_width_mode = st.selectbox("PDF column width mode",
        ["Auto (by content)", "Equal", "Custom (inches, comma-separated)"])
    custom_widths_input = ""
    if pdf_width_mode == "Custom (inches, comma-separated)":
        custom_widths_input = st.text_input("Custom widths (e.g. 1.0,2.5,1.5)")

    pdf_period = (f"{date_range[0].strftime('%d %b %Y')} – {date_range[1].strftime('%d %b %Y')}"
                  if isinstance(date_range, tuple) and len(date_range) == 2 else str(date_range))

    if REPORTLAB_AVAILABLE:
        def df_to_pdf_bytes(df_in, title="Bank submit status", subtitle="", custom_widths=None, generated_at=""):
            buf = BytesIO()
            lm = rm = tm = bm = 24
            from reportlab.lib.pagesizes import landscape as _landscape2
            pw, ph = _landscape2(A3)
            uw = pw - lm - rm
            hf = "Helvetica-Bold"; cf = "Helvetica"; hfs = 10; cfs = 8
            min_cw = 1.2*inch; max_cw = 4.5*inch

            def mw(txt, font, sz): return pdfmetrics.stringWidth(str(txt), font, sz)

            col_widths_pdf = []
            if custom_widths and isinstance(custom_widths, (list, tuple)):
                col_widths_pdf = [max(min_cw, min(max_cw, float(w)*inch)) for w in custom_widths]
                if len(col_widths_pdf) < len(df_in.columns):
                    rem = len(df_in.columns) - len(col_widths_pdf)
                    add = max(0, uw - sum(col_widths_pdf)) / rem if rem else min_cw
                    col_widths_pdf.extend([max(min_cw, min(max_cw, add))]*rem)
                tot = sum(col_widths_pdf)
                if tot > uw and tot > 0: col_widths_pdf = [w*uw/tot for w in col_widths_pdf]
            elif custom_widths == "EQUAL":
                col_widths_pdf = [uw/len(df_in.columns)]*len(df_in.columns)
            else:
                dt = df_in.fillna("").astype(str)
                for col in df_in.columns:
                    hw = mw(col, hf, hfs)
                    vals = dt[col].tolist()
                    if len(vals) > 250: vals = vals[::max(1, len(vals)//250)]
                    measured = sorted([mw(v, cf, cfs) for v in vals if v])
                    cw2 = measured[min(len(measured)-1, int(len(measured)*0.9))] if measured else mw("M", cf, cfs)
                    col_widths_pdf.append(min(max_cw, max(min_cw, max(hw, cw2)+16)))
                tot = sum(col_widths_pdf)
                if tot > uw and tot > 0: col_widths_pdf = [w*uw/tot for w in col_widths_pdf]
                elif tot < uw and tot > 0: col_widths_pdf = [w + (uw-tot)*(w/tot) for w in col_widths_pdf]

            ss2 = getSampleStyleSheet()
            hs  = ParagraphStyle("H", parent=ss2["Normal"], fontName=hf, fontSize=hfs,
                                 leading=11, textColor=colors.white, alignment=TA_LEFT, wordWrap="CJK")
            cs  = ParagraphStyle("C", parent=ss2["Normal"], fontName=cf, fontSize=cfs,
                                 leading=10, alignment=TA_LEFT, wordWrap="CJK")

            def chunk(cols, widths, max_w):
                groups, cur, cw3 = [], [], 0.0
                for c, w in zip(cols, widths):
                    if cur and cw3 + w > max_w: groups.append(cur); cur=[c]; cw3=w
                    else: cur.append(c); cw3+=w
                if cur: groups.append(cur)
                return groups

            ts = TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0d3f47")),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("ALIGN",(0,0),(-1,-1),"LEFT"), ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("FONTNAME",(0,0),(-1,-1),"Helvetica"), ("FONTSIZE",(0,0),(-1,-1),8),
                ("LEFTPADDING",(0,0),(-1,-1),5), ("RIGHTPADDING",(0,0),(-1,-1),5),
                ("BOTTOMPADDING",(0,0),(-1,-1),4), ("TOPPADDING",(0,0),(-1,-1),4),
                ("GRID",(0,0),(-1,-1),0.25,colors.grey), ("BOX",(0,0),(-1,-1),0.5,colors.black),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.whitesmoke,colors.lightgrey]),
            ])
            pages = []
            for g_i, g_cols in enumerate(chunk(list(df_in.columns), col_widths_pdf, uw)):
                g_w = [col_widths_pdf[list(df_in.columns).index(c)] for c in g_cols]
                rows = [[Paragraph(str(c), hs) for c in g_cols]]
                for rv in df_in.fillna("").astype(str).values.tolist():
                    rows.append([Paragraph(str(rv[list(df_in.columns).index(c)]), cs) for c in g_cols])
                tbl2 = LongTable(rows, repeatRows=1, colWidths=g_w, hAlign="LEFT", splitByRow=1, spaceBefore=12, spaceAfter=12)
                tbl2.setStyle(ts)
                pages.append(tbl2)
                if g_i < len(chunk(list(df_in.columns), col_widths_pdf, uw))-1:
                    pages.append(PageBreak())

            def ph2(canvas, doc):
                canvas.saveState()
                ty=ph-tm+10; sy=ty-16
                canvas.setFont(hf,20); canvas.drawCentredString(pw/2,ty,str(title))
                canvas.setFont(cf,11); canvas.drawCentredString(pw/2,sy,str(subtitle))
                if generated_at: canvas.setFont(cf,8); canvas.drawString(lm,sy-14,f"Generated on: {generated_at}")
                canvas.setFont(cf,8); canvas.drawRightString(pw-rm,ty,f"Page {doc.page}")
                canvas.setFont(hf,8); canvas.drawRightString(pw-rm,bm/2,"ASM")
                canvas.restoreState()

            doc = SimpleDocTemplate(buf, pagesize=_landscape(A3), leftMargin=lm, rightMargin=rm, topMargin=tm+28, bottomMargin=bm)
            doc.build(pages, onFirstPage=ph2, onLaterPages=ph2)
            buf.seek(0); return buf.read()

        try:
            parsed_custom = None
            if pdf_width_mode == "Auto (by content)":
                parsed_custom = [max(1.2, min(4.5, col_widths.get(c, 100)/96.0)) for c in pft_display.columns]
            elif pdf_width_mode == "Equal":
                parsed_custom = "EQUAL"
            elif pdf_width_mode == "Custom (inches, comma-separated)" and custom_widths_input:
                try: parsed_custom = [float(x.strip()) for x in custom_widths_input.split(",") if x.strip()]
                except Exception: col_pdf.error("Invalid custom widths.")

            pdf_df = pft_export.copy()
            if "Invoice Value" in pdf_df.columns:
                pdf_df["Invoice Value"] = pdf_df["Invoice Value"].map(lambda x: f"${x:,.2f}")
            if len(pdf_df.columns) > 0:
                tr = {c: "" for c in pdf_df.columns}
                tr[pdf_df.columns[0]] = "TOTAL"
                if "Invoice Value" in pdf_df.columns:
                    tr["Invoice Value"] = f"${total_invoice_value:,.2f}"
                pdf_df = pd.concat([pdf_df, pd.DataFrame([tr])], ignore_index=True)

            pdf_bytes = df_to_pdf_bytes(pdf_df, subtitle=pdf_period,
                custom_widths=parsed_custom, generated_at=datetime.now().strftime("%d %b %Y %H:%M:%S"))
            col_pdf.download_button("📄 Download PDF", pdf_bytes, "bank_submit_status.pdf", "application/pdf")
        except Exception as e:
            col_pdf.error(f"Could not generate PDF: {e}")
    else:
        col_pdf.warning("Install reportlab: pip install reportlab")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — BANK ACCEPT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t_accept:
    st.markdown("## ✅ Bank Accept Analysis")
    df["Status"] = df.apply(lambda r:
        "Paid" if pd.notna(r["Payment. Rcv Dt"]) else
        "Accepted" if pd.notna(r["Bank Accept Date"]) else
        "Not Accepted", axis=1)

    total_acc  = df[df["Status"]=="Accepted"]["Invoice Value"].sum()
    total_nacc = df[df["Status"]=="Not Accepted"]["Invoice Value"].sum()
    total_paid = df[df["Status"]=="Paid"]["Invoice Value"].sum()
    total_value = df["Invoice Value"].sum()
    acceptance_rate = (total_acc / total_value * 100) if total_value > 0 else 0
    total_bills = len(df)
    acc_bills = len(df[df["Status"]=="Accepted"])
    nacc_bills = len(df[df["Status"]=="Not Accepted"])
    paid_bills = len(df[df["Status"]=="Paid"])

    ba_kpi_css = """
    <style>
    .ba-3d-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 14px;
        margin: 14px 0;
    }
    .ba-3d-card {
        background: linear-gradient(145deg, #112240, #0a1628);
        border: 1px solid rgba(0,201,167,0.18);
        border-radius: 14px;
        padding: 22px 18px;
        transform: perspective(700px) rotateX(3deg) rotateY(0deg) translateZ(0);
        box-shadow: 0 8px 25px rgba(0,0,0,0.45), 0 4px 10px rgba(0,201,167,0.06),
                    inset 0 1px 0 rgba(255,255,255,0.06), inset 0 -2px 0 rgba(0,0,0,0.3);
        transition: all 0.25s cubic-bezier(.4,0,.2,1);
        position: relative;
        overflow: hidden;
    }
    .ba-3d-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
        background: linear-gradient(90deg, transparent, var(--card-accent, rgba(0,201,167,0.5)), transparent);
    }
    .ba-3d-card::after {
        content: '';
        position: absolute;
        bottom: 0; left: 10%; right: 10%;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,201,167,0.2), transparent);
    }
    .ba-3d-card:hover {
        transform: perspective(700px) rotateX(0deg) rotateY(0deg) translateZ(8px) translateY(-4px);
        box-shadow: 0 16px 40px rgba(0,201,167,0.15), 0 8px 20px rgba(0,0,0,0.5),
                    inset 0 1px 0 rgba(255,255,255,0.10);
        border-color: rgba(0,201,167,0.45);
    }
    .ba-3d-card:active {
        transform: perspective(700px) rotateX(2deg) translateZ(2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.4), inset 0 2px 4px rgba(0,0,0,0.2);
    }
    .ba-3d-icon { font-size: 20px; margin-bottom: 6px; }
    .ba-3d-label { font-size: 11px; color: #8899aa; font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; margin-bottom: 6px; }
    .ba-3d-value { font-size: 28px; font-weight: 800; line-height: 1.1; text-shadow: 0 1px 3px rgba(0,0,0,0.4); }
    .ba-3d-value.green { color: #00c9a7; }
    .ba-3d-value.red { color: #ff3b30; }
    .ba-3d-value.gold { color: #ffcc00; }
    .ba-3d-delta { font-size: 11px; color: #8899aa; margin-top: 6px; font-weight: 500; }
    </style>
    """
    st.markdown(ba_kpi_css, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="ba-3d-grid">
        <div class="ba-3d-card" style="--card-accent: rgba(0,201,167,0.6);">
            <div class="ba-3d-icon">✅</div>
            <div class="ba-3d-label">Accepted Value</div>
            <div class="ba-3d-value green">{usd(total_acc)}</div>
            <div class="ba-3d-delta">{acc_bills:,} bills</div>
        </div>
        <div class="ba-3d-card" style="--card-accent: rgba(255,59,48,0.6);">
            <div class="ba-3d-icon">❌</div>
            <div class="ba-3d-label">Not Accepted Value</div>
            <div class="ba-3d-value red">{usd(total_nacc)}</div>
            <div class="ba-3d-delta">{nacc_bills:,} bills</div>
        </div>
        <div class="ba-3d-card" style="--card-accent: rgba(26,143,255,0.6);">
            <div class="ba-3d-icon">💰</div>
            <div class="ba-3d-label">Paid Value</div>
            <div class="ba-3d-value green">{usd(total_paid)}</div>
            <div class="ba-3d-delta">{paid_bills:,} bills</div>
        </div>
        <div class="ba-3d-card" style="--card-accent: rgba(255,204,0,0.6);">
            <div class="ba-3d-icon">📈</div>
            <div class="ba-3d-label">Acceptance Rate</div>
            <div class="ba-3d-value gold">{acceptance_rate:.1f}%</div>
            <div class="ba-3d-delta">{total_bills:,} total bills</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### 📊 Status Distribution")
    fig_acc = px.pie(df, names="Status", values="Invoice Value",
                     color_discrete_sequence=["#00c9a7", "#1a8fff", "#ff6b35"], hole=0.45)
    fig_acc.update_layout(**PL, height=350)
    st.plotly_chart(fig_acc, use_container_width=True)
    st.markdown("---")

    st.markdown("## 📅 Bank Accept Date-wise Analysis")
    df_accepted = df[df["Status"]=="Accepted"].copy()

    if not df_accepted.empty:
        st.markdown("### 📈 Daily Accepted Value Trend")
        fig_trend = go.Figure()
        top_firms_acc = (df_accepted.groupby("Firm Name")["Invoice Value"].sum()
                         .sort_values(ascending=False).head(10).index.tolist())
        df_accepted["Firm Group"] = df_accepted["Firm Name"].apply(
            lambda x: x if x in top_firms_acc else "Others")
        for firm in top_firms_acc + ["Others"]:
            firm_data = df_accepted[df_accepted["Firm Group"] == firm]
            if not firm_data.empty:
                daily_firm = (firm_data.groupby(firm_data["Bank Accept Date"].dt.date)
                              .agg(Value=("Invoice Value","sum")).reset_index().sort_values("Bank Accept Date"))
                fig_trend.add_scatter(x=daily_firm["Bank Accept Date"], y=daily_firm["Value"],
                                      mode="lines+markers", name=firm, line=dict(width=2), marker=dict(size=5))
        fig_trend.update_layout(**PL_GENERAL,
            xaxis=dict(title="Bank Accept Date", tickangle=-45, gridcolor="#1a2a3a"),
            yaxis=dict(title="Accepted Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            legend=dict(orientation="h", y=1.1, x=0, bgcolor="rgba(0,0,0,0)", font=dict(color="#8899aa", size=9)),
            height=400, hovermode="x unified")
        st.plotly_chart(fig_trend, use_container_width=True)
        st.markdown("---")

        st.markdown("### 📋 Date-wise Accepted Summary")
        pivot_table = (df_accepted.pivot_table(
            index=df_accepted["Bank Accept Date"].dt.date,
            columns="Firm Name", values="Invoice Value", aggfunc="sum", fill_value=0))
        pivot_table["Total"] = pivot_table.sum(axis=1)
        display_df2 = pivot_table.copy()
        for col in display_df2.columns:
            display_df2[col] = display_df2[col].map(lambda x: f"${x:,.2f}")
        st.dataframe(display_df2, use_container_width=True, height=400)

        st.markdown("---")
        st.markdown("### 🏢 Firm-wise Detailed Breakdown")
        firm_summary = (df_accepted.groupby("Firm Name")
                        .agg(TotalAcceptedValue=("Invoice Value","sum"),
                             TotalAcceptedCount=("LC No","count"),
                             FirstAcceptDate=("Bank Accept Date","min"),
                             LastAcceptDate=("Bank Accept Date","max"),
                             AvgValuePerAcceptance=("Invoice Value","mean"))
                        .reset_index().sort_values("TotalAcceptedValue", ascending=False))
        firm_summary["TotalAcceptedValue"] = firm_summary["TotalAcceptedValue"].map(lambda x: f"${x:,.2f}")
        firm_summary["AvgValuePerAcceptance"] = firm_summary["AvgValuePerAcceptance"].map(lambda x: f"${x:,.2f}")
        firm_summary["FirstAcceptDate"] = firm_summary["FirstAcceptDate"].dt.strftime("%d %b %Y")
        firm_summary["LastAcceptDate"]  = firm_summary["LastAcceptDate"].dt.strftime("%d %b %Y")
        st.dataframe(firm_summary, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ No accepted records found for the current filters.")
    st.markdown("---")
    st.caption(f"Showing acceptance data for {len(df_accepted):,} accepted records out of {len(df):,} total submissions")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — ASM ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with t_asm:
    st.markdown("## 📊 ASM Analysis")
    df["Status"] = df.apply(lambda r:
        "Paid" if pd.notna(r["Payment. Rcv Dt"]) else
        "Accepted" if pd.notna(r["Bank Accept Date"]) else
        "Not Accepted", axis=1)

    total_acc2  = df[df["Status"]=="Accepted"]["Invoice Value"].sum()
    total_nacc2 = df[df["Status"]=="Not Accepted"]["Invoice Value"].sum()
    total_paid2 = df[df["Status"]=="Paid"]["Invoice Value"].sum()
    total_count2 = len(df)
    total_value2 = df["Invoice Value"].sum()
    acceptance_rate2 = (total_acc2 / total_value2 * 100) if total_value2 > 0 else 0

    total_acc_count2 = len(df[df["Status"]=="Accepted"])
    total_nacc_count2 = len(df[df["Status"]=="Not Accepted"])
    total_paid_count2 = len(df[df["Status"]=="Paid"])

    asm_kpi_css = """
    <style>
    .asm-3d-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 14px;
        margin: 14px 0;
    }
    .asm-3d-card {
        background: linear-gradient(145deg, #112240, #0a1628);
        border: 1px solid rgba(0,201,167,0.18);
        border-radius: 14px;
        padding: 20px 18px;
        transform: perspective(700px) rotateX(3deg) rotateY(0deg) translateZ(0);
        box-shadow: 0 8px 25px rgba(0,0,0,0.45), 0 4px 10px rgba(0,201,167,0.06),
                    inset 0 1px 0 rgba(255,255,255,0.06), inset 0 -2px 0 rgba(0,0,0,0.3);
        transition: all 0.25s cubic-bezier(.4,0,.2,1);
        position: relative;
        overflow: hidden;
    }
    .asm-3d-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
        background: linear-gradient(90deg, transparent, var(--card-accent, rgba(0,201,167,0.5)), transparent);
    }
    .asm-3d-card::after {
        content: '';
        position: absolute;
        bottom: 0; left: 10%; right: 10%;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0,201,167,0.2), transparent);
    }
    .asm-3d-card:hover {
        transform: perspective(700px) rotateX(0deg) rotateY(0deg) translateZ(8px) translateY(-4px);
        box-shadow: 0 16px 40px rgba(0,201,167,0.15), 0 8px 20px rgba(0,0,0,0.5),
                    inset 0 1px 0 rgba(255,255,255,0.10);
        border-color: rgba(0,201,167,0.45);
    }
    .asm-3d-card:active {
        transform: perspective(700px) rotateX(2deg) translateZ(2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.4), inset 0 2px 4px rgba(0,0,0,0.2);
    }
    .asm-3d-icon { font-size: 20px; margin-bottom: 6px; }
    .asm-3d-label { font-size: 11px; color: #8899aa; font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; margin-bottom: 6px; }
    .asm-3d-value { font-size: 26px; font-weight: 800; color: #00c9a7; line-height: 1.1; text-shadow: 0 1px 3px rgba(0,0,0,0.4); }
    .asm-3d-delta { font-size: 11px; color: #8899aa; margin-top: 6px; font-weight: 500; }
    .asm-3d-value.green { color: #00c9a7; }
    .asm-3d-value.red { color: #ff3b30; }
    .asm-3d-value.blue { color: #1a8fff; }
    .asm-3d-value.gold { color: #ffcc00; }
    </style>
    """
    st.markdown(asm_kpi_css, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="asm-3d-grid" style="grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));">
        <div class="asm-3d-card">
            <div class="asm-3d-icon">📋</div>
            <div class="asm-3d-label">Total Bill Count</div>
            <div class="asm-3d-value">{total_count2:,}</div>
            <div class="asm-3d-delta">total records</div>
        </div>
        <div class="asm-3d-card">
            <div class="asm-3d-icon">✅</div>
            <div class="asm-3d-label">Accepted Value</div>
            <div class="asm-3d-value green">{usd(total_acc2)}</div>
            <div class="asm-3d-delta">{total_acc_count2:,} bills accepted</div>
        </div>
        <div class="asm-3d-card">
            <div class="asm-3d-icon">❌</div>
            <div class="asm-3d-label">Not Accepted Value</div>
            <div class="asm-3d-value red">{usd(total_nacc2)}</div>
            <div class="asm-3d-delta">{total_nacc_count2:,} bills not accepted</div>
        </div>
        <div class="asm-3d-card">
            <div class="asm-3d-icon">💰</div>
            <div class="asm-3d-label">Paid Value</div>
            <div class="asm-3d-value blue">{usd(total_paid2)}</div>
            <div class="asm-3d-delta">{total_paid_count2:,} bills paid</div>
        </div>
        <div class="asm-3d-card">
            <div class="asm-3d-icon">📈</div>
            <div class="asm-3d-label">Overall Acceptance Rate</div>
            <div class="asm-3d-value gold">{acceptance_rate2:.1f}%</div>
            <div class="asm-3d-delta">of total value</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    fig_acc2 = px.pie(df, names="Status", values="Invoice Value",
                      color_discrete_sequence=["#00c9a7","#1a8fff","#ff6b35"], hole=0.45)
    fig_acc2.update_layout(**PL)
    st.plotly_chart(fig_acc2, use_container_width=True)
    st.markdown("---")

    st.markdown("### 📅 Daily Acceptance Trend")
    df_acc2 = df[df["Status"]=="Accepted"].copy()
    if not df_acc2.empty:
        acc_trend2 = (df_acc2.groupby(df_acc2["Bank Accept Date"].dt.date)
                     .agg(AcceptedValue=("Invoice Value","sum"), AcceptedCount=("LC No","count"))
                     .reset_index().sort_values("Bank Accept Date"))
        fig_acc_trend2 = go.Figure()
        fig_acc_trend2.add_scatter(x=acc_trend2["Bank Accept Date"], y=acc_trend2["AcceptedValue"],
                                   mode="lines+markers", line=dict(color="#00c9a7", width=2.5), marker=dict(size=6))
        fig_acc_trend2.update_layout(**PL_GENERAL,
            xaxis=dict(title="Bank Accept Date", tickangle=-40, gridcolor="#1a2a3a"),
            yaxis=dict(title="Accepted Value (USD)", gridcolor="#1a2a3a"),
            height=360, showlegend=False)
        st.plotly_chart(fig_acc_trend2, use_container_width=True)
        acc_trend2["AcceptedValue"] = acc_trend2["AcceptedValue"].map(lambda x: f"${x:,.2f}")
        st.dataframe(acc_trend2, use_container_width=True, hide_index=True)
    else:
        st.info("No accepted records found for the current filters.")
    st.markdown("---")

    st.markdown("### 📊 Bank Accept Date Firm-wise Analysis")
    last_acc_date = df["Bank Accept Date"].dropna().max()
    if not pd.isna(last_acc_date):
        df_latest = df[df["Bank Accept Date"].dt.date == last_acc_date.date()].copy()
        total_acc_latest  = df_latest[df_latest["Status"]=="Accepted"]["Invoice Value"].sum()
        total_nacc_latest = df_latest[df_latest["Status"]=="Not Accepted"]["Invoice Value"].sum()
        total_records_latest = len(df_latest)

        c1b, c2b, c3b, c4b = st.columns(4)

        baf_kpi_css = """
        <style>
        .baf-3d-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 14px;
            margin: 14px 0;
        }
        .baf-3d-card {
            background: linear-gradient(145deg, #112240, #0a1628);
            border: 1px solid rgba(0,201,167,0.18);
            border-radius: 14px;
            padding: 20px 18px;
            transform: perspective(700px) rotateX(3deg) rotateY(0deg) translateZ(0);
            box-shadow: 0 8px 25px rgba(0,0,0,0.45), 0 4px 10px rgba(0,201,167,0.06),
                        inset 0 1px 0 rgba(255,255,255,0.06), inset 0 -2px 0 rgba(0,0,0,0.3);
            transition: all 0.25s cubic-bezier(.4,0,.2,1);
            position: relative;
            overflow: hidden;
        }
        .baf-3d-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 2px;
            background: linear-gradient(90deg, transparent, var(--card-accent, rgba(0,201,167,0.5)), transparent);
        }
        .baf-3d-card::after {
            content: '';
            position: absolute;
            bottom: 0; left: 10%; right: 10%;
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(0,201,167,0.2), transparent);
        }
        .baf-3d-card:hover {
            transform: perspective(700px) rotateX(0deg) rotateY(0deg) translateZ(8px) translateY(-4px);
            box-shadow: 0 16px 40px rgba(0,201,167,0.15), 0 8px 20px rgba(0,0,0,0.5),
                        inset 0 1px 0 rgba(255,255,255,0.10);
            border-color: rgba(0,201,167,0.45);
        }
        .baf-3d-card:active {
            transform: perspective(700px) rotateX(2deg) translateZ(2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.4), inset 0 2px 4px rgba(0,0,0,0.2);
        }
        .baf-3d-icon { font-size: 20px; margin-bottom: 6px; }
        .baf-3d-label { font-size: 11px; color: #8899aa; font-weight: 600; letter-spacing: 0.04em; text-transform: uppercase; margin-bottom: 6px; }
        .baf-3d-value { font-size: 26px; font-weight: 800; line-height: 1.1; text-shadow: 0 1px 3px rgba(0,0,0,0.4); }
        .baf-3d-value.teal { color: #00c9a7; }
        .baf-3d-value.green { color: #00c9a7; }
        .baf-3d-value.red { color: #ff3b30; }
        .baf-3d-value.white { color: #e2eaf3; }
        .baf-3d-delta { font-size: 11px; color: #8899aa; margin-top: 6px; font-weight: 500; }
        </style>
        """
        st.markdown(baf_kpi_css, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="baf-3d-grid">
            <div class="baf-3d-card" style="--card-accent: rgba(26,143,255,0.6);">
                <div class="baf-3d-icon">📅</div>
                <div class="baf-3d-label">Latest Accept Date</div>
                <div class="baf-3d-value white">{last_acc_date.strftime("%d %b %Y")}</div>
            </div>
            <div class="baf-3d-card" style="--card-accent: rgba(0,201,167,0.6);">
                <div class="baf-3d-icon">✅</div>
                <div class="baf-3d-label">Accepted Value</div>
                <div class="baf-3d-value green">{usd(total_acc_latest)}</div>
            </div>
            <div class="baf-3d-card" style="--card-accent: rgba(255,59,48,0.6);">
                <div class="baf-3d-icon">❌</div>
                <div class="baf-3d-label">Not Accepted Value</div>
                <div class="baf-3d-value red">{usd(total_nacc_latest)}</div>
            </div>
            <div class="baf-3d-card" style="--card-accent: rgba(0,201,167,0.6);">
                <div class="baf-3d-icon">📋</div>
                <div class="baf-3d-label">Total Records</div>
                <div class="baf-3d-value teal">{total_records_latest:,}</div>
                <div class="baf-3d-delta">bills on latest date</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        firm_acceptance = (df_latest[df_latest["Status"]=="Accepted"]
                           .groupby("Firm Name")
                           .agg(AcceptedCount=("LC No","count"), AcceptedValue=("Invoice Value","sum"))
                           .reset_index().sort_values("AcceptedValue", ascending=False))
        total_accepted_count = firm_acceptance["AcceptedCount"].sum()
        total_accepted_value = firm_acceptance["AcceptedValue"].sum()
        firm_acceptance["Percentage"] = (firm_acceptance["AcceptedValue"] / total_accepted_value * 100).round(1)

        display_table = firm_acceptance.copy()
        display_table["AcceptedValue"] = display_table["AcceptedValue"].map(lambda x: f"${x:,.2f}")
        display_table["Percentage"]    = display_table["Percentage"].map(lambda x: f"{x:.1f}%")
        display_table.columns = ["Firm Name","Accepted Count","Accepted Value (USD)","% of Total"]
        total_row_asm = {"Firm Name":"**TOTAL**","Accepted Count":total_accepted_count,
                     "Accepted Value (USD)":f"**${total_accepted_value:,.2f}**","% of Total":"**100%**"}
        display_table = pd.concat([display_table, pd.DataFrame([total_row_asm])], ignore_index=True)
        st.dataframe(display_table, use_container_width=True, hide_index=True)
        st.markdown("---")
        st.caption(f"📊 Latest Bank Accept Date: {last_acc_date.strftime('%d %b %Y')} | Total {total_records_latest:,} Records")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 9 — DUE DATE TRACKER 🗓️
# ══════════════════════════════════════════════════════════════════════════════
with t_due:
    st.markdown("## 🗓️ Due Date Tracker")
    st.caption(f"Today: **{today.strftime('%d %b %Y')}**")

    if "Maturity Date" not in df.columns:
        st.warning("Maturity Date column not found in data.")
    else:
        df_dd = df[df["Payment. Rcv Dt"].isna() & df["Maturity Date"].notna()].copy()
        df_dd["Days Until Maturity"] = (today - df_dd["Maturity Date"]).dt.days
        df_dd["Due Status"] = df_dd["Days Until Maturity"].apply(
            lambda d: "🔴 Overdue"      if d > 0
            else      "🟡 Due in 7d"    if d >= -7
            else      "🟠 Due in 15d"   if d >= -15
            else      "🔵 Due in 30d"   if d >= -30
            else      "🟢 Due in 60d+"  if d >= -60
            else      "✅ Safe"
        )

        # Summary KPIs
        overdue_df = df_dd[df_dd["Days Until Maturity"] > 0]
        due7_df    = df_dd[(df_dd["Days Until Maturity"] <= 0) & (df_dd["Days Until Maturity"] >= -7)]
        due15_df   = df_dd[(df_dd["Days Until Maturity"] < -7) & (df_dd["Days Until Maturity"] >= -15)]
        due30_df   = df_dd[(df_dd["Days Until Maturity"] < -15) & (df_dd["Days Until Maturity"] >= -30)]
        due60_df   = df_dd[(df_dd["Days Until Maturity"] < -30) & (df_dd["Days Until Maturity"] >= -60)]

        kpi_data = [
            ("🔴", "Overdue",    len(overdue_df), overdue_df['Invoice Value'].sum(), "#ff3b30"),
            ("🟡", "Due in 7d",  len(due7_df),    due7_df['Invoice Value'].sum(),    "#ff9500"),
            ("🟠", "Due in 15d", len(due15_df),   due15_df['Invoice Value'].sum(),   "#ff6b35"),
            ("🔵", "Due in 30d", len(due30_df),   due30_df['Invoice Value'].sum(),   "#1a8fff"),
            ("🟢", "Due in 60d", len(due60_df),   due60_df['Invoice Value'].sum(),   "#00c9a7"),
        ]

        kpi_html = '<div style="display:flex; gap:12px; flex-wrap:wrap; justify-content:space-between; margin:12px 0;">'
        for icon, label, count, value, color in kpi_data:
            kpi_html += f'''
            <div style="flex:1; min-width:160px;
                        background: linear-gradient(145deg, #1a2235, #111827);
                        border: 1px solid {color}33;
                        border-radius: 16px;
                        padding: 20px 16px;
                        box-shadow: 6px 6px 16px rgba(0,0,0,0.5),
                                    -4px -4px 12px rgba(255,255,255,0.03),
                                    inset 0 1px 0 rgba(255,255,255,0.05);
                        transform: perspective(600px) rotateX(2deg) rotateY(-1deg);
                        transition: transform 0.3s ease, box-shadow 0.3s ease;
                        text-align:center;">
                <div style="font-size:28px; margin-bottom:6px;
                            filter: drop-shadow(0 2px 4px {color}44);">{icon}</div>
                <div style="font-size:13px; color:#8899aa; font-weight:600; letter-spacing:0.5px;
                            text-transform:uppercase; margin-bottom:4px;">{label}</div>
                <div style="font-size:28px; font-weight:800; color:#e2eaf3;
                            text-shadow: 0 2px 8px rgba(0,0,0,0.4),
                                         0 0 20px {color}22;
                            margin-bottom:4px;">{count} <span style="font-size:14px; font-weight:500; color:{color};">Bill</span></div>
                <div style="font-size:13px; color:{color}; font-weight:600;
                            background: {color}15;
                            border-radius:20px; padding:3px 10px; display:inline-block;
                            box-shadow: inset 0 1px 3px rgba(0,0,0,0.2);">↑ {usd(value)}</div>
            </div>'''
        kpi_html += '</div>'
        st.markdown(kpi_html, unsafe_allow_html=True)
        st.markdown("---")

        # Filter by period
        period_filter = st.radio("Show Bill's due in:",
            ["All Unpaid", "Overdue", "Next 7 Days", "Next 15 Days", "Next 30 Days", "Next 60 Days"],
            horizontal=True, key="due_period")

        if period_filter == "Overdue":
            df_show = overdue_df
        elif period_filter == "Next 7 Days":
            df_show = df_dd[df_dd["Days Until Maturity"].between(-7, 0)]
        elif period_filter == "Next 15 Days":
            df_show = df_dd[df_dd["Days Until Maturity"].between(-15, 0)]
        elif period_filter == "Next 30 Days":
            df_show = df_dd[df_dd["Days Until Maturity"].between(-30, 0)]
        elif period_filter == "Next 60 Days":
            df_show = df_dd[df_dd["Days Until Maturity"].between(-60, 0)]
        else:
            df_show = df_dd

        df_show = df_show.sort_values("Days Until Maturity", ascending=False)

        search_due = st.text_input("🔍 Search Firm / Bill / Party / Bank", key="due_search")
        if search_due:
            mask = df_show.astype(str).apply(
                lambda c: c.str.contains(search_due, case=False, na=False)).any(axis=1)
            df_show = df_show[mask]

        cols_due = ["Firm Name","LC No","Bank Refno","Party Name","Our Bank","Bank Name","Invoice Value",
                    "Maturity Date","Days Until Maturity","Due Status","Sales Person","Payment. Rcv Dt"]
        cols_due = [c for c in cols_due if c in df_show.columns]
        df_due_display = df_show[cols_due].copy()
        df_due_display["Maturity Date"] = df_due_display["Maturity Date"].dt.strftime("%d %b %Y")
        if "Payment. Rcv Dt" in df_due_display.columns:
            df_due_display["Payment. Rcv Dt"] = pd.to_datetime(df_due_display["Payment. Rcv Dt"], errors="coerce").dt.strftime("%d %b %Y").fillna("")
        df_due_display["Invoice Value"] = df_due_display["Invoice Value"].map(lambda x: f"${x:,.2f}")
        df_due_display["Days Until Maturity"] = df_due_display["Days Until Maturity"].apply(lambda x: f"{int(x)}")

        def color_days(val):
            try:
                v = int(val)
                if v > 0:    return "background-color:#ff3b3022; color:#ff3b30; font-weight:700;"
                elif v >= -7: return "background-color:#ff950018; color:#ff9500; font-weight:700;"
                elif v >= -15: return "background-color:#ff6b3518; color:#ff6b35; font-weight:600;"
                elif v >= -30: return "background-color:#1a8fff15; color:#1a8fff; font-weight:600;"
                else:          return "background-color:#00c9a715; color:#00c9a7;"
            except: return ""

        def color_status(val):
            s = str(val)
            if "Overdue" in s:     return "color:#ff3b30; font-weight:700;"
            elif "7d" in s:        return "color:#ff9500; font-weight:700;"
            elif "15d" in s:       return "color:#ff6b35; font-weight:600;"
            elif "30d" in s:       return "color:#1a8fff; font-weight:600;"
            else:                  return "color:#00c9a7;"

        styled = df_due_display.style
        if "Days Until Maturity" in df_due_display.columns:
            styled = styled.map(color_days, subset=["Days Until Maturity"])
        if "Due Status" in df_due_display.columns:
            styled = styled.map(color_status, subset=["Due Status"])

        st.dataframe(styled, use_container_width=True, hide_index=True, height=min(500, 38*len(df_due_display)+50))
        st.caption(f"Showing **{len(df_due_display):,}** of **{len(df_dd):,}** records | Total unpaid value: **{usd(df_show['Invoice Value'].sum())}**")

        # ─── 1. Firm-wise Overdue Summary (Top 15) ─────────────────────────
        st.markdown("---")
        sh("1. Firm-wise Overdue Summary (Top 15)")

        _overdue_only = df_show[df_show["Days Until Maturity"] > 0]
        if len(_overdue_only) > 0:
            firm_sum = (_overdue_only.groupby("Firm Name")
                       .agg(Bill_Count=("Invoice Value", "size"),
                            Total_Bill_Value=("Invoice Value", "sum"),
                            Max_Overdue=("Days Until Maturity", "max"))
                       .sort_values("Total_Bill_Value", ascending=False).head(15).reset_index())
            firm_sum["Risk_Level"] = firm_sum["Max_Overdue"].apply(
                lambda d: "CRITICAL" if d > 60 else "HIGH" if d > 30 else "MEDIUM" if d > 14 else "LOW")
            firm_sum["#"] = range(1, len(firm_sum) + 1)
            firm_sum_raw = firm_sum.copy()
            firm_sum["Total_Bill_Value"] = firm_sum["Total_Bill_Value"].map(lambda x: f"${x/1000:.1f}K" if x >= 1000 else f"${x:.0f}")
            firm_sum["Max_Overdue"] = firm_sum["Max_Overdue"].map(lambda x: f"{int(x)}d")
            firm_display = firm_sum[["#", "Firm Name", "Bill_Count", "Total_Bill_Value", "Max_Overdue", "Risk_Level"]].copy()
            firm_display.columns = ["#", "Firm Name", "Bill Count", "Total Bill Value", "Max Overdue", "Risk Level"]

            def style_risk_firm(val):
                s = str(val)
                if s == "CRITICAL": return "color:#ff3b30; font-weight:800; background:#ff3b3018; border-radius:6px; padding:2px 8px;"
                elif s == "HIGH":   return "color:#ff9500; font-weight:700; background:#ff950018; border-radius:6px; padding:2px 8px;"
                elif s == "MEDIUM": return "color:#ff6b35; font-weight:600; background:#ff6b3518; border-radius:6px; padding:2px 8px;"
                else:               return "color:#00c9a7; font-weight:600; background:#00c9a718; border-radius:6px; padding:2px 8px;"
            st.dataframe(firm_display.style.map(style_risk_firm, subset=["Risk Level"]),
                         use_container_width=True, hide_index=True, height=min(500, 38*len(firm_display)+38))

            firm_chart = (_overdue_only.groupby("Firm Name")
                         .agg(Value=("Invoice Value", "sum"))
                         .sort_values("Value", ascending=True).tail(15))
            fig_firm = go.Figure(go.Bar(
                x=firm_chart["Value"].values, y=firm_chart.index, orientation="h",
                marker=dict(color=firm_chart["Value"].values, colorscale="RdYlGn_r"),
                text=[usd(v) for v in firm_chart["Value"].values], textposition="outside"))
            fig_firm.update_layout(**PL_GENERAL,
                height=max(280, len(firm_chart)*30),
                xaxis=dict(title="Invoice Value (USD)", gridcolor="#1a2a3a"),
                yaxis=dict(gridcolor="#1a2a3a"), showlegend=False)
            st.plotly_chart(fig_firm, use_container_width=True)

        # ─── Firm Drill-down — Click firm name → popup window ─────────────
            st.markdown('<div class="firm-card-section">', unsafe_allow_html=True)
            firm_names_list = firm_sum_raw["Firm Name"].tolist()
            st.markdown("---")
            st.markdown("##### 👆 Click any firm name to open detail popup")
            st.session_state["_firm_overdue_data"] = _overdue_only.copy()

            firm_btn_css = """
            <style>
            .firm-card-container {
                position: relative;
                margin-bottom: 8px;
            }
            .firm-3d-card {
                background: linear-gradient(145deg, #112240, #0a1628);
                border: 1px solid rgba(0,201,167,0.20);
                border-radius: 14px;
                padding: 22px 24px;
                cursor: pointer;
                transform: perspective(600px) rotateX(4deg) rotateY(0deg) translateZ(0);
                box-shadow: 0 8px 25px rgba(0,0,0,0.45), 0 4px 10px rgba(0,201,167,0.08),
                            inset 0 1px 0 rgba(255,255,255,0.06), inset 0 -2px 0 rgba(0,0,0,0.3);
                transition: all 0.25s cubic-bezier(.4,0,.2,1);
                position: relative;
                overflow: hidden;
            }
            .firm-3d-card::before {
                content: '';
                position: absolute;
                top: 0; left: 0; right: 0;
                height: 2px;
                background: linear-gradient(90deg, transparent, rgba(0,201,167,0.5), transparent);
            }
            .firm-3d-card::after {
                content: '';
                position: absolute;
                bottom: 0; left: 10%; right: 10%;
                height: 1px;
                background: linear-gradient(90deg, transparent, rgba(0,201,167,0.3), transparent);
            }
            .firm-3d-card:hover {
                transform: perspective(600px) rotateX(0deg) rotateY(0deg) translateZ(8px) translateY(-4px);
                box-shadow: 0 16px 40px rgba(0,201,167,0.18), 0 8px 20px rgba(0,0,0,0.5),
                            inset 0 1px 0 rgba(255,255,255,0.10);
                border-color: rgba(0,201,167,0.45);
            }
            .firm-3d-name {
                font-size: 18px;
                font-weight: 800;
                color: #e2eaf3;
                margin-bottom: 8px;
                text-shadow: 0 1px 3px rgba(0,0,0,0.4);
            }
            .firm-3d-meta {
                font-size: 12px;
                color: #8899aa;
                display: flex;
                gap: 10px;
                flex-wrap: wrap;
            }
            .firm-3d-meta span {
                background: rgba(0,201,167,0.08);
                border: 1px solid rgba(0,201,167,0.15);
                border-radius: 8px;
                padding: 3px 8px;
                font-weight: 600;
            }
            .firm-3d-overdue { color: #ff3b30; }
            .firm-click-hint {
                font-size: 10px;
                color: rgba(0,201,167,0.5);
                text-align: center;
                margin-top: 8px;
                letter-spacing: 0.05em;
                text-transform: uppercase;
            }
            .firm-card-section div[data-testid="stColumn"] .stButton {
                position: relative;
                margin-top: -72px;
                margin-bottom: 62px;
                z-index: 5;
            }
            .firm-card-section div[data-testid="stColumn"] .stButton > button {
                background: transparent !important;
                color: transparent !important;
                border: none !important;
                box-shadow: none !important;
                width: 100% !important;
                height: 72px !important;
                min-height: 72px !important;
                padding: 0 !important;
                cursor: pointer !important;
                border-radius: 14px !important;
                transition: all 0.2s ease !important;
            }
            .firm-card-section div[data-testid="stColumn"] .stButton > button:hover {
                background: rgba(0,201,167,0.06) !important;
                box-shadow: 0 0 20px rgba(0,201,167,0.12) !important;
            }
            .firm-card-section div[data-testid="stColumn"] .stButton > button:active {
                transform: scale(0.98) !important;
            }
            </style>
            """
            st.markdown(firm_btn_css, unsafe_allow_html=True)

            cols_per_row = 3
            for i in range(0, len(firm_names_list), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, col in enumerate(cols):
                    idx = i + j
                    if idx < len(firm_names_list):
                        fname = firm_names_list[idx]
                        row = firm_sum_raw[firm_sum_raw["Firm Name"] == fname].iloc[0]
                        rc = {"CRITICAL":"#ff3b30","HIGH":"#ff9500","MEDIUM":"#ff6b35","LOW":"#00c9a7"}.get(row["Risk_Level"],"#00c9a7")
                        with col:
                            st.markdown(f"""
                            <div class="firm-card-container">
                                <div class="firm-3d-card" style="border-left: 3px solid {rc};">
                                    <div class="firm-3d-name">{fname}</div>
                                    <div class="firm-3d-meta">
                                        <span>{row['Bill_Count']} Bills</span>
                                        <span>{usd(row['Total_Bill_Value'])}</span>
                                        <span class="firm-3d-overdue">{int(row['Max_Overdue'])}d overdue</span>
                                    </div>
                                    <div class="firm-click-hint">▸ Click to view details</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            if st.button("​", key=f"firm_open_{idx}", use_container_width=True):
                                show_firm_detail(fname)
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("No overdue bills in current filter.")

        # ─── 2. Our Bank-wise Overdue Summary ───────────────────────────────
        st.markdown("---")
        sh("2. Our Bank-wise Overdue Summary")

        if len(_overdue_only) > 0:
            bank_sum = (_overdue_only.groupby("Our Bank")
                       .agg(Bill_Count=("Invoice Value", "size"),
                            Total_Value=("Invoice Value", "sum"),
                            Max_Days=("Days Until Maturity", "max"))
                       .sort_values("Total_Value", ascending=False).reset_index())
            bank_sum["Risk_Level"] = bank_sum["Max_Days"].apply(
                lambda d: "CRITICAL" if d > 60 else "HIGH" if d > 30 else "MEDIUM" if d > 14 else "LOW")
            bank_sum["#"] = range(1, len(bank_sum) + 1)
            bank_sum["Total_Value"] = bank_sum["Total_Value"].map(lambda x: f"${x/1000:.1f}K" if x >= 1000 else f"${x:.0f}")
            bank_sum["Max_Days"] = bank_sum["Max_Days"].map(lambda x: f"{int(x)}d")
            bank_display = bank_sum[["#", "Our Bank", "Bill_Count", "Total_Value", "Max_Days", "Risk_Level"]].copy()
            bank_display.columns = ["#", "Our Bank", "Bill Count", "Total Value", "Max Days", "Risk Level"]

            def style_risk_bank(val):
                s = str(val)
                if s == "CRITICAL": return "color:#ff3b30; font-weight:800; background:#ff3b3018; border-radius:6px; padding:2px 8px;"
                elif s == "HIGH":   return "color:#ff9500; font-weight:700; background:#ff950018; border-radius:6px; padding:2px 8px;"
                elif s == "MEDIUM": return "color:#ff6b35; font-weight:600; background:#ff6b3518; border-radius:6px; padding:2px 8px;"
                else:               return "color:#00c9a7; font-weight:600; background:#00c9a718; border-radius:6px; padding:2px 8px;"
            st.dataframe(bank_display.style.map(style_risk_bank, subset=["Risk Level"]),
                         use_container_width=True, hide_index=True, height=min(400, 38*len(bank_display)+38))

            bank_chart = (_overdue_only.groupby("Our Bank")
                         .agg(Value=("Invoice Value", "sum"))
                         .sort_values("Value", ascending=True))
            fig_bank = go.Figure(go.Bar(
                x=bank_chart["Value"].values, y=bank_chart.index, orientation="h",
                marker=dict(color=bank_chart["Value"].values, colorscale="RdYlGn_r"),
                text=[usd(v) for v in bank_chart["Value"].values], textposition="outside"))
            fig_bank.update_layout(**PL_GENERAL,
                height=max(250, len(bank_chart)*30),
                xaxis=dict(title="Invoice Value (USD)", gridcolor="#1a2a3a"),
                yaxis=dict(gridcolor="#1a2a3a"), showlegend=False)
            st.plotly_chart(fig_bank, use_container_width=True)

            # Pie chart for bank distribution
            bank_pie = _overdue_only.groupby("Our Bank").agg(Value=("Invoice Value","sum")).reset_index()
            fig_bank_pie = px.pie(bank_pie, names="Our Bank", values="Value",
                                   color_discrete_sequence=C, hole=0.4)
            fig_bank_pie.update_layout(**PL_GENERAL, height=320,
                legend=dict(font=dict(color="#8899aa", size=11)))
            st.plotly_chart(fig_bank_pie, use_container_width=True)
        else:
            st.info("No overdue bills in current filter.")

        # ─── 3. Sales Person-wise Overdue Summary ───────────────────────────
        st.markdown("---")
        sh("3. Sales Person-wise Overdue Summary")

        if len(_overdue_only) > 0:
            sp_data = _overdue_only.copy()
            sp_data["Sales Person"] = sp_data["Sales Person"].fillna("(No Sales Person)")
            sp_sum = (sp_data.groupby("Sales Person")
                     .agg(Bill_Count=("Invoice Value", "size"),
                          Total_Value=("Invoice Value", "sum"),
                          Max_Days=("Days Until Maturity", "max"),
                          Avg_Days=("Days Until Maturity", "mean"))
                     .sort_values("Total_Value", ascending=False).reset_index())
            sp_sum["Risk_Level"] = sp_sum["Max_Days"].apply(
                lambda d: "CRITICAL" if d > 60 else "HIGH" if d > 30 else "MEDIUM" if d > 14 else "LOW")
            sp_sum["#"] = range(1, len(sp_sum) + 1)
            sp_sum["Total_Value"] = sp_sum["Total_Value"].map(lambda x: f"${x/1000:.1f}K" if x >= 1000 else f"${x:.0f}")
            sp_sum["Max_Days"] = sp_sum["Max_Days"].map(lambda x: f"{int(x)}d")
            sp_sum["Avg_Days"] = sp_sum["Avg_Days"].map(lambda x: f"{x:.0f}d")
            sp_display = sp_sum[["#", "Sales Person", "Bill_Count", "Total_Value", "Max_Days", "Avg_Days", "Risk_Level"]].copy()
            sp_display.columns = ["#", "Sales Person", "Bill Count", "Total Value", "Max Days", "Avg Days", "Risk Level"]

            def style_risk_sp(val):
                s = str(val)
                if s == "CRITICAL": return "color:#ff3b30; font-weight:800; background:#ff3b3018; border-radius:6px; padding:2px 8px;"
                elif s == "HIGH":   return "color:#ff9500; font-weight:700; background:#ff950018; border-radius:6px; padding:2px 8px;"
                elif s == "MEDIUM": return "color:#ff6b35; font-weight:600; background:#ff6b3518; border-radius:6px; padding:2px 8px;"
                else:               return "color:#00c9a7; font-weight:600; background:#00c9a718; border-radius:6px; padding:2px 8px;"
            st.dataframe(sp_display.style.map(style_risk_sp, subset=["Risk Level"]),
                         use_container_width=True, hide_index=True, height=min(500, 38*len(sp_display)+38))

            sp_chart = (sp_data.groupby("Sales Person")
                       .agg(Value=("Invoice Value", "sum"), Count=("Invoice Value", "size"))
                       .sort_values("Value", ascending=True))
            fig_sp = go.Figure(go.Bar(
                x=sp_chart["Value"].values, y=sp_chart.index, orientation="h",
                marker=dict(color=sp_chart["Value"].values, colorscale="RdYlGn_r"),
                text=[f"{usd(v)} ({c})" for v, c in zip(sp_chart["Value"].values, sp_chart["Count"].values)],
                textposition="outside"))
            fig_sp.update_layout(**PL_GENERAL,
                height=max(250, len(sp_chart)*35),
                xaxis=dict(title="Invoice Value (USD)", gridcolor="#1a2a3a"),
                yaxis=dict(gridcolor="#1a2a3a"), showlegend=False)
            st.plotly_chart(fig_sp, use_container_width=True)

            # Pie chart for sales person distribution
            sp_pie = sp_data.groupby("Sales Person").agg(Value=("Invoice Value","sum")).reset_index()
            fig_sp_pie = px.pie(sp_pie, names="Sales Person", values="Value",
                                 color_discrete_sequence=C, hole=0.4)
            fig_sp_pie.update_layout(**PL_GENERAL, height=320,
                legend=dict(font=dict(color="#8899aa", size=11)))
            st.plotly_chart(fig_sp_pie, use_container_width=True)
        else:
            st.info("No overdue bills in current filter.")

        # Chart: Maturity timeline
        st.markdown("---")
        sh("📊 Maturity Value by Due Status")
        due_summary = (df_dd.groupby("Due Status")
                       .agg(Count=("LC No","count"), Value=("Invoice Value","sum"))
                       .reset_index().sort_values("Value", ascending=False))
        fig_due = px.bar(due_summary, x="Due Status", y="Value", color="Due Status",
                         color_discrete_sequence=["#ff3b30","#ff9500","#ff6b35","#1a8fff","#00c9a7","#44dd66"],
                         text=due_summary["Value"].map(usd))
        fig_due.update_traces(textposition="outside", textfont_size=11)
        fig_due.update_layout(**PL_GENERAL,
            xaxis=dict(title="", gridcolor="#1a2a3a"),
            yaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            showlegend=False, height=350)
        st.plotly_chart(fig_due, use_container_width=True)

        # Download
        col_csv, col_pdf_due = st.columns(2)
        col_csv.download_button("📥 Download Due Date List (CSV)",
            df_due_display.to_csv(index=False).encode("utf-8"),
            "due_date_tracker.csv", "text/csv")

        if REPORTLAB_AVAILABLE:
            def due_date_report_pdf(df_in, period_label, generated_at, summary_data):
                buf = BytesIO()
                from reportlab.lib.pagesizes import landscape as _landscape2
                from reportlab.platypus import Spacer
                from reportlab.lib.enums import TA_RIGHT
                pw, ph = _landscape2(A3)
                lm = rm = 40; tm = 50; bm = 45
                uw = pw - lm - rm

                hf = "Helvetica-Bold"; cf = "Helvetica"
                title_clr = colors.HexColor("#0d3f47")
                accent    = colors.HexColor("#00c9a7")
                dark_bg   = colors.HexColor("#0b1220")
                mid_bg    = colors.HexColor("#1a2235")
                header_bg = colors.HexColor("#0d3f47")
                header_fg = colors.white
                text_dark = colors.HexColor("#222222")
                text_mut  = colors.HexColor("#666666")
                overdue_c = colors.HexColor("#ff3b30")
                warn_c    = colors.HexColor("#ff9500")
                orange_c  = colors.HexColor("#ff6b35")
                info_c    = colors.HexColor("#1a8fff")
                safe_c    = colors.HexColor("#00c9a7")
                row_even  = colors.HexColor("#f5f7fa")
                row_odd   = colors.white
                border_c  = colors.HexColor("#d0d5dd")

                def mw(txt, font, sz):
                    return pdfmetrics.stringWidth(str(txt), font, sz)

                def _usd_func(v):
                    try: v = float(v)
                    except: return "$0.00"
                    if v >= 1e6: return "$" + f"{v/1e6:.2f}" + "M"
                    if v >= 1e3: return "$" + f"{v/1e3:.1f}" + "K"
                    return "$" + f"{v:.2f}"

                _filtered_overdue = df_show[df_show["Days Until Maturity"] > 0]

                _firm_overdue = _filtered_overdue.groupby("Firm Name").agg(
                    count=("Invoice Value", "size"),
                    total_value=("Invoice Value", "sum"),
                    max_overdue_days=("Days Until Maturity", "max"),
                ).sort_values("total_value", ascending=False).head(15) if len(_filtered_overdue) > 0 else pd.DataFrame()

                _bank_overdue = _filtered_overdue.groupby("Our Bank").agg(
                    count=("Invoice Value", "size"),
                    total_value=("Invoice Value", "sum"),
                    max_overdue_days=("Days Until Maturity", "max"),
                    avg_overdue_days=("Days Until Maturity", "mean"),
                ).sort_values("total_value", ascending=False).head(15) if len(_filtered_overdue) > 0 else pd.DataFrame()

                # ═══════════════════════════════════════════════════════════
                # PAGE 1 — COVER
                # ═══════════════════════════════════════════════════════════
                def draw_cover(canvas, doc):
                    canvas.saveState()
                    canvas.setFillColor(dark_bg); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
                    canvas.setFillColor(accent); canvas.roundRect(pw/2 - 200, ph - 180, 400, 8, 4, fill=1, stroke=0)
                    canvas.setFillColor(header_fg); canvas.setFont(hf, 32)
                    canvas.drawCentredString(pw/2, ph - 250, "DUE DATE TRACKER")
                    canvas.setFillColor(accent); canvas.setFont(hf, 18)
                    canvas.drawCentredString(pw/2, ph - 285, "Bank Submit History Report")
                    canvas.setFillColor(colors.HexColor("#8899aa")); canvas.setFont(cf, 12)
                    canvas.drawCentredString(pw/2, ph - 320, f"Report Period: {period_label}")
                    canvas.drawCentredString(pw/2, ph - 340, f"Generated: {generated_at}")
                    canvas.setFillColor(accent); canvas.roundRect(pw/2 - 200, ph - 370, 400, 2, 1, fill=1, stroke=0)

                    box_y = ph - 530; box_h = 110; box_w = (uw - 40) / 4
                    kpi_boxes = [
                        ("TOTAL RECORDS", str(summary_data['total']), summary_data['total_val'], info_c),
                        ("TOTAL VALUE", summary_data['total_val'], str(summary_data['total']) + " Bills", accent),
                        ("OVERDUE", str(summary_data['overdue']) + " Bills", summary_data['overdue_val'], overdue_c),
                        ("DUE SOON (7d)", str(summary_data['due7']) + " Bills", summary_data['due7_val'], warn_c),
                    ]
                    for i, (lbl, val, sub, clr) in enumerate(kpi_boxes):
                        bx = lm + 20 + i * (box_w + 10)
                        canvas.setFillColor(mid_bg); canvas.roundRect(bx, box_y, box_w, box_h, 8, fill=1, stroke=0)
                        canvas.setStrokeColor(clr); canvas.setLineWidth(3)
                        canvas.line(bx, box_y + box_h, bx + box_w, box_y + box_h)
                        canvas.setFillColor(clr); canvas.setFont(hf, 10); canvas.drawString(bx + 15, box_y + box_h - 25, lbl)
                        canvas.setFillColor(header_fg); canvas.setFont(hf, 26); canvas.drawString(bx + 15, box_y + 40, val)
                        canvas.setFillColor(clr); canvas.setFont(hf, 14); canvas.drawString(bx + 15, box_y + 12, sub)

                    canvas.setFillColor(colors.HexColor("#556677")); canvas.setFont(cf, 9)
                    canvas.drawCentredString(pw/2, 100, "Prepared by: Smart Dashboard v2.0")
                    canvas.drawCentredString(pw/2, 82, "Confidential - For Internal Use Only")
                    canvas.setStrokeColor(accent); canvas.setLineWidth(0.5); canvas.line(lm, 60, pw - rm, 60)
                    canvas.restoreState()

                # ═══════════════════════════════════════════════════════════
                # PAGE 2 — EXECUTIVE SUMMARY
                # ═══════════════════════════════════════════════════════════
                def draw_summary(canvas, doc):
                    canvas.saveState()
                    canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 20)
                    canvas.drawString(lm, ph - tm - 10, "EXECUTIVE SUMMARY")
                    canvas.setStrokeColor(accent); canvas.setLineWidth(2)
                    canvas.line(lm, ph - tm - 18, lm + 220, ph - tm - 18)
                    canvas.setFillColor(text_mut); canvas.setFont(cf, 10)
                    canvas.drawString(lm, ph - tm - 38, f"Report Period: {period_label}  |  Generated: {generated_at}")

                    sec_y = ph - tm - 70
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
                    canvas.drawString(lm, sec_y, "1. Maturity Status Overview")
                    canvas.setStrokeColor(border_c); canvas.setLineWidth(0.5)
                    canvas.line(lm, sec_y - 5, pw - rm, sec_y - 5)

                    card_y = sec_y - 95; card_h = 75; card_w = (uw - 30) / 5
                    cards = [
                        ("Overdue", summary_data['overdue'], summary_data['overdue_val'], overdue_c),
                        ("Due in 7 Days", summary_data['due7'], summary_data['due7_val'], warn_c),
                        ("Due in 15 Days", summary_data['due15'], summary_data['due15_val'], orange_c),
                        ("Due in 30 Days", summary_data['due30'], summary_data['due30_val'], info_c),
                        ("Due in 60 Days", summary_data['due60'], summary_data['due60_val'], safe_c),
                    ]
                    for i, (lbl, cnt, val, clr) in enumerate(cards):
                        cx = lm + i * (card_w + 6)
                        canvas.setFillColor(colors.HexColor("#f8f9fa")); canvas.roundRect(cx, card_y, card_w, card_h, 6, fill=1, stroke=0)
                        canvas.setFillColor(clr); canvas.roundRect(cx, card_y + card_h - 5, card_w, 5, 2, fill=1, stroke=0)
                        canvas.setFont(hf, 9); canvas.drawString(cx + 10, card_y + card_h - 22, lbl)
                        canvas.setFillColor(text_dark); canvas.setFont(hf, 22); canvas.drawString(cx + 10, card_y + 30, str(cnt) + " Bills")
                        canvas.setFillColor(clr); canvas.setFont(hf, 11); canvas.drawString(cx + 10, card_y + 8, val)

                    sec2_y = card_y - 40
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
                    canvas.drawString(lm, sec2_y, "2. Firm-wise Overdue Summary (Top 15)")
                    canvas.setStrokeColor(border_c); canvas.line(lm, sec2_y - 5, pw - rm, sec2_y - 5)

                    if len(_firm_overdue) > 0:
                        firm_header = ["#", "Firm Name", "Bill Count", "Total Bill Value", "Max Overdue", "Risk Level"]
                        firm_rows = [firm_header]
                        for rank, (firm, frow) in enumerate(_firm_overdue.iterrows(), 1):
                            max_d = int(frow["max_overdue_days"])
                            cnt = int(frow["count"])
                            val = _usd_func(frow["total_value"])
                            risk = "CRITICAL" if max_d > 60 else "HIGH" if max_d > 30 else "MEDIUM" if max_d > 15 else "LOW"
                            firm_rows.append([str(rank), str(firm)[:30], str(cnt), val, str(max_d) + "d", risk])

                        fcol_w = [25, 180, 60, 100, 65, 70]
                        ftbl = TableStyle([
                            ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                            ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
                            ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
                            ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                            ("ALIGN", (1, 0), (1, -1), "LEFT"),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                            ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                            ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
                        ])
                        for ri in range(1, len(firm_rows)):
                            rv2 = firm_rows[ri][-1]
                            if rv2 == "CRITICAL": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); ftbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                            elif rv2 == "HIGH": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); ftbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                            elif rv2 == "MEDIUM": ftbl.add("TEXTCOLOR", (-1, ri), (-1, ri), orange_c)
                        ft = LongTable(firm_rows, colWidths=fcol_w, hAlign="LEFT")
                        ft.setStyle(ftbl)
                        tw, th = ft.wrap(520, 200)
                        ft.drawOn(canvas, lm, sec2_y - 25 - th)
                    else:
                        th = 0

                    bank_x = lm + 560
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
                    canvas.drawString(bank_x, sec2_y, "3. Our Bank-wise Overdue Summary")
                    canvas.setStrokeColor(border_c); canvas.line(bank_x, sec2_y - 5, bank_x + 520, sec2_y - 5)

                    if len(_bank_overdue) > 0:
                        bheader = ["#", "Our Bank", "Bill Count", "Total Value", "Max Days", "Risk Level"]
                        brows = [bheader]
                        for rank, (bname, brow) in enumerate(_bank_overdue.iterrows(), 1):
                            max_d = int(brow["max_overdue_days"])
                            cnt = int(brow["count"])
                            val = _usd_func(brow["total_value"])
                            risk = "CRITICAL" if max_d > 60 else "HIGH" if max_d > 30 else "MEDIUM" if max_d > 15 else "LOW"
                            brows.append([str(rank), str(bname)[:25], str(cnt), val, str(max_d) + "d", risk])

                        bcol_w = [25, 120, 55, 85, 55, 70]
                        btbl = TableStyle([
                            ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                            ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
                            ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
                            ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                            ("ALIGN", (1, 0), (1, -1), "LEFT"),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                            ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                            ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
                        ])
                        for ri in range(1, len(brows)):
                            rv2 = brows[ri][-1]
                            if rv2 == "CRITICAL": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); btbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                            elif rv2 == "HIGH": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); btbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                            elif rv2 == "MEDIUM": btbl.add("TEXTCOLOR", (-1, ri), (-1, ri), orange_c)
                        bt = LongTable(brows, colWidths=bcol_w, hAlign="LEFT")
                        bt.setStyle(btbl)
                        btw, bth = bt.wrap(420, 200)
                        bt.drawOn(canvas, bank_x, sec2_y - 25 - bth)
                    else:
                        bth = 0

                    sec3_y = sec2_y - 25 - max(th, bth) - 35
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
                    canvas.drawString(lm, sec3_y, "4. Key Observations")
                    canvas.setStrokeColor(border_c); canvas.line(lm, sec3_y - 5, pw - rm, sec3_y - 5)

                    total_overdue_val = overdue_df['Invoice Value'].sum()
                    total_all_val = df_show['Invoice Value'].sum()
                    pct_overdue = (total_overdue_val / total_all_val * 100) if total_all_val > 0 else 0
                    top_firm = _firm_overdue.index[0] if len(_firm_overdue) > 0 else "N/A"
                    top_firm_val = _usd_func(_firm_overdue.iloc[0]["total_value"]) if len(_firm_overdue) > 0 else "$0"
                    top_bank = _bank_overdue.index[0] if len(_bank_overdue) > 0 else "N/A"
                    top_bank_val = _usd_func(_bank_overdue.iloc[0]["total_value"]) if len(_bank_overdue) > 0 else "$0"
                    critical_firms = len(_firm_overdue[_firm_overdue["max_overdue_days"] > 60]) if len(_firm_overdue) > 0 else 0

                    obs = [
                        f"OVERALL RISK: {summary_data['overdue']} bills ({pct_overdue:.1f}% of total value) are overdue, totaling {summary_data['overdue_val']}.",
                        f"TOP FIRM: {top_firm} leads with {top_firm_val} overdue. Immediate collection follow-up recommended.",
                        f"TOP BANK: {top_bank} exposure is {top_bank_val}. Coordinate with relationship manager.",
                        f"CRITICAL: {critical_firms} firms overdue >60 days. Escalate to senior management for recovery.",
                    ]
                    oy = sec3_y - 20
                    for i, o in enumerate(obs):
                        canvas.setFillColor(overdue_c if i == 0 else warn_c if i < 3 else info_c)
                        canvas.setFont(hf if i < 2 else cf, 8)
                        canvas.drawString(lm + 5, oy - i * 15, o)

                    canvas.setFillColor(text_mut); canvas.setFont(cf, 8)
                    canvas.drawString(lm, bm - 15, f"Page {doc.page}  |  Bank Submit History Report")
                    canvas.drawRightString(pw - rm, bm - 15, "Confidential")
                    canvas.restoreState()

                # ═══════════════════════════════════════════════════════════
                # PAGE 3 — OVERDUE RECOVERY SUGGESTIONS
                # ═══════════════════════════════════════════════════════════
                def draw_suggestions(canvas, doc):
                    canvas.saveState()
                    canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 20)
                    canvas.drawString(lm, ph - tm - 10, "OVERDUE RECOVERY SUGGESTIONS")
                    canvas.setStrokeColor(accent); canvas.setLineWidth(2)
                    canvas.line(lm, ph - tm - 18, lm + 300, ph - tm - 18)
                    canvas.setFillColor(text_mut); canvas.setFont(cf, 10)
                    canvas.drawString(lm, ph - tm - 38, "Firm-wise & Bank-wise action plan for CO/ED review  |  " + generated_at)

                    sec_y = ph - tm - 75
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
                    canvas.drawString(lm, sec_y, "1. Firm-wise Overdue Recovery Roadmap")
                    canvas.setStrokeColor(border_c); canvas.setLineWidth(0.5)
                    canvas.line(lm, sec_y - 5, pw - rm, sec_y - 5)

                    sug_header = ["#", "Firm Name", "Bills", "Bill Value", "Max Days", "Suggested Action", "Priority"]
                    sug_rows = [sug_header]
                    if len(_firm_overdue) > 0:
                        for rank, (firm, frow) in enumerate(_firm_overdue.head(12).iterrows(), 1):
                            cnt = int(frow["count"]); max_d = int(frow["max_overdue_days"])
                            val = _usd_func(frow["total_value"])
                            if max_d > 60: action = "Immediate escalation to CO. Legal notice."; priority = "URGENT"
                            elif max_d > 30 and cnt > 3: action = "Schedule meeting with firm accounts dept."; priority = "HIGH"
                            elif max_d > 30: action = "Send formal reminder. Follow up weekly."; priority = "MEDIUM"
                            elif cnt > 5: action = "Bulk collection drive via sales team."; priority = "MEDIUM"
                            else: action = "Standard follow-up via sales team."; priority = "LOW"
                            sug_rows.append([str(rank), str(firm)[:30], str(cnt), val, str(max_d) + "d", action, priority])

                    scol_w = [22, 150, 40, 85, 50, 310, 60]
                    stbl = TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                        ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
                        ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
                        ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (4, -1), "RIGHT"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
                    ])
                    for ri in range(1, len(sug_rows)):
                        pv = sug_rows[ri][-1]
                        if pv == "URGENT": stbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); stbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                        elif pv == "HIGH": stbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); stbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                    st = LongTable(sug_rows, colWidths=scol_w, hAlign="LEFT")
                    st.setStyle(stbl)
                    stw, sth = st.wrap(uw, 200)
                    st.drawOn(canvas, lm, sec_y - 25 - sth)

                    sec_b_y = sec_y - 25 - sth - 30
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 13)
                    canvas.drawString(lm, sec_b_y, "2. Our Bank-wise Overdue Recovery Roadmap")
                    canvas.setStrokeColor(border_c); canvas.line(lm, sec_b_y - 5, pw - rm, sec_b_y - 5)

                    bsug_header = ["#", "Our Bank", "Bills", "Overdue Value", "Max Days", "Bank Action Required", "Priority"]
                    bsug_rows = [bsug_header]
                    if len(_bank_overdue) > 0:
                        for rank, (bname, brow) in enumerate(_bank_overdue.head(10).iterrows(), 1):
                            max_d = int(brow["max_overdue_days"]); cnt = int(brow["count"]); val = _usd_func(brow["total_value"])
                            if max_d > 60: action = "Escalate to bank management. Request written update."; priority = "URGENT"
                            elif max_d > 30: action = "Meet relationship manager. Push for clearance."; priority = "HIGH"
                            else: action = "Regular follow-up with bank contact."; priority = "MEDIUM"
                            bsug_rows.append([str(rank), str(bname)[:25], str(cnt), val, str(max_d) + "d", action, priority])

                    bsc_w = [22, 100, 40, 85, 50, 355, 60]
                    bstbl = TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                        ("FONTNAME", (0, 0), (-1, 0), hf), ("FONTSIZE", (0, 0), (-1, 0), 8),
                        ("FONTNAME", (0, 1), (-1, -1), cf), ("FONTSIZE", (0, 1), (-1, -1), 7),
                        ("ALIGN", (0, 0), (0, -1), "CENTER"), ("ALIGN", (2, 0), (4, -1), "RIGHT"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
                    ])
                    for ri in range(1, len(bsug_rows)):
                        pv = bsug_rows[ri][-1]
                        if pv == "URGENT": bstbl.add("TEXTCOLOR", (-1, ri), (-1, ri), overdue_c); bstbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                        elif pv == "HIGH": bstbl.add("TEXTCOLOR", (-1, ri), (-1, ri), warn_c); bstbl.add("FONTNAME", (-1, ri), (-1, ri), hf)
                    bst = LongTable(bsug_rows, colWidths=bsc_w, hAlign="LEFT")
                    bst.setStyle(bstbl)
                    bstw, bsth = bst.wrap(uw, 200)
                    bst.drawOn(canvas, lm, sec_b_y - 25 - bsth)

                    box_top = sec_b_y - 25 - bsth - 25
                    canvas.setFillColor(colors.HexColor("#f0f9ff")); canvas.roundRect(lm, box_top - 160, uw, 160, 8, fill=1, stroke=0)
                    canvas.setStrokeColor(info_c); canvas.setLineWidth(1.5); canvas.roundRect(lm, box_top - 160, uw, 160, 8, fill=0, stroke=1)
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 12)
                    canvas.drawString(lm + 15, box_top - 20, "3. General Recovery Strategy")
                    suggestions = [
                        "1. PRIORITY MATRIX: Classify overdue bills into Critical (>60d), High (>30d), Medium (>15d) and Low risk.",
                        "2. FIRM ENGAGEMENT: Schedule direct meetings with top overdue firms. Prepare firm-wise outstanding statements.",
                        "3. BANK COORDINATION: Work with relationship managers at CBP, SEBPLC, DBBL for faster clearance.",
                        "4. SALES TEAM ALERT: Assign specific sales persons to follow up on their respective firm overdue.",
                        "5. WEEKLY REVIEW: Establish weekly overdue review meeting with CO to track recovery progress.",
                    ]
                    sy = box_top - 38
                    for i, s in enumerate(suggestions):
                        canvas.setFillColor(text_dark); canvas.setFont(cf, 8)
                        canvas.drawString(lm + 20, sy - i * 18, s)

                    canvas.setFillColor(text_mut); canvas.setFont(cf, 8)
                    canvas.drawString(lm, bm - 15, f"Page {doc.page}  |  Bank Submit History Report")
                    canvas.drawRightString(pw - rm, bm - 15, "Confidential")
                    canvas.restoreState()

                # ═══════════════════════════════════════════════════════════
                # PAGE 4+ — DATA TABLE
                # ═══════════════════════════════════════════════════════════
                def draw_data_page(canvas, doc):
                    canvas.saveState()
                    canvas.setFillColor(colors.white); canvas.rect(0, 0, pw, ph, fill=1, stroke=0)
                    canvas.setFillColor(title_clr); canvas.setFont(hf, 16)
                    canvas.drawString(lm, ph - tm - 5, "DETAILED BILL DATA")
                    canvas.setStrokeColor(accent); canvas.setLineWidth(1.5)
                    canvas.line(lm, ph - tm - 12, lm + 180, ph - tm - 12)
                    canvas.setFillColor(text_mut); canvas.setFont(cf, 9)
                    canvas.drawRightString(pw - rm, ph - tm - 8, f"Page {doc.page}  |  {period_label}")
                    canvas.setStrokeColor(border_c); canvas.setLineWidth(0.3)
                    canvas.line(lm, bm + 5, pw - rm, bm + 5)
                    canvas.setFillColor(text_mut); canvas.setFont(cf, 7)
                    canvas.drawString(lm, bm - 8, f"Smart Dashboard v2.0  |  Generated: {generated_at}")
                    canvas.drawRightString(pw - rm, bm - 8, "Confidential")
                    canvas.restoreState()

                elements = []
                elements.append(Spacer(1, 20)); elements.append(PageBreak())  # p2 summary
                elements.append(Spacer(1, 20)); elements.append(PageBreak())  # p3 suggestions
                elements.append(Spacer(1, 20)); elements.append(PageBreak())  # p4 data

                dt = df_in.fillna("").astype(str)
                col_list = list(df_in.columns)
                col_widths_pdf = []
                for col in col_list:
                    vals = dt[col].tolist()
                    if len(vals) > 100: vals = vals[::max(1, len(vals)//100)]
                    measured = [mw(v, cf, 7.5) for v in vals if v]
                    mx = max([mw(col, hf, 9)] + measured) if measured else mw(col, hf, 9)
                    col_widths_pdf.append(max(55, min(200, mx + 12)))
                tot_w = sum(col_widths_pdf)
                if tot_w > uw: col_widths_pdf = [w * uw / tot_w for w in col_widths_pdf]

                def chunk(cols, widths, max_w):
                    groups, cur, cw3 = [], [], 0.0
                    for c, w in zip(cols, widths):
                        if cur and cw3 + w > max_w: groups.append(cur); cur = [c]; cw3 = w
                        else: cur.append(c); cw3 += w
                    if cur: groups.append(cur)
                    return groups

                ss2 = getSampleStyleSheet()
                hs = ParagraphStyle("RH", parent=ss2["Normal"], fontName=hf, fontSize=8, leading=10, textColor=header_fg, alignment=TA_LEFT)
                cs = ParagraphStyle("RC", parent=ss2["Normal"], fontName=cf, fontSize=7, leading=9, alignment=TA_LEFT)

                _pg_callbacks = {2: draw_summary, 3: draw_suggestions}
                def _draw_later(canvas, doc):
                    fn = _pg_callbacks.get(doc.page, draw_data_page)
                    fn(canvas, doc)

                groups = chunk(col_list, col_widths_pdf, uw)
                for g_i, g_cols in enumerate(groups):
                    g_w = [col_widths_pdf[col_list.index(c)] for c in g_cols]
                    rows_data = [[Paragraph(str(c), hs) for c in g_cols]]
                    for rv in df_in.fillna("").astype(str).values.tolist():
                        row_cells = []
                        for c in g_cols:
                            idx = col_list.index(c)
                            val = str(rv[idx])
                            style = cs
                            if c == "Due Status":
                                if "Overdue" in val:  style = ParagraphStyle("OVR"+str(g_i), parent=cs, textColor=overdue_c, fontName=hf)
                                elif "7d" in val:      style = ParagraphStyle("D7"+str(g_i), parent=cs, textColor=warn_c, fontName=hf)
                                elif "15d" in val:     style = ParagraphStyle("D15"+str(g_i), parent=cs, textColor=orange_c)
                                elif "30d" in val:     style = ParagraphStyle("D30"+str(g_i), parent=cs, textColor=info_c)
                                elif "60d" in val:     style = ParagraphStyle("D60"+str(g_i), parent=cs, textColor=safe_c)
                            elif c == "Days Until Maturity":
                                try:
                                    dv = int(val)
                                    if dv > 0:    style = ParagraphStyle("DVp"+str(g_i), parent=cs, textColor=overdue_c, fontName=hf)
                                    elif dv >= -7: style = ParagraphStyle("DV7"+str(g_i), parent=cs, textColor=warn_c, fontName=hf)
                                    elif dv >= -15: style = ParagraphStyle("DV15"+str(g_i), parent=cs, textColor=orange_c)
                                    elif dv >= -30: style = ParagraphStyle("DV30"+str(g_i), parent=cs, textColor=info_c)
                                    else:          style = ParagraphStyle("DV60"+str(g_i), parent=cs, textColor=safe_c)
                                except: pass
                            row_cells.append(Paragraph(val, style))
                        rows_data.append(row_cells)

                    tbl = LongTable(rows_data, repeatRows=1, colWidths=g_w, hAlign="LEFT", splitByRow=1, spaceBefore=6, spaceAfter=6)
                    ts = TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), header_bg), ("TEXTCOLOR", (0, 0), (-1, 0), header_fg),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 7),
                        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 2),
                        ("GRID", (0, 0), (-1, -1), 0.3, border_c), ("BOX", (0, 0), (-1, -1), 0.6, header_bg),
                        ("ROWBACKGROUNDS", (1, 1), (-1, -1), [row_even, row_odd]),
                    ])
                    tbl.setStyle(ts)
                    elements.append(tbl)
                    if g_i < len(groups) - 1: elements.append(PageBreak())

                doc = SimpleDocTemplate(buf, pagesize=_landscape2(A3), leftMargin=lm, rightMargin=rm, topMargin=tm + 20, bottomMargin=bm)
                doc.build(elements, onFirstPage=draw_cover, onLaterPages=_draw_later)
                buf.seek(0)
                return buf.read()

            # Build summary data from filtered df_show
            _f_overdue = df_show[df_show["Days Until Maturity"] > 0]
            _f_due7    = df_show[(df_show["Days Until Maturity"] <= 0) & (df_show["Days Until Maturity"] >= -7)]
            _f_due15   = df_show[(df_show["Days Until Maturity"] < -7)  & (df_show["Days Until Maturity"] >= -15)]
            _f_due30   = df_show[(df_show["Days Until Maturity"] < -15) & (df_show["Days Until Maturity"] >= -30)]
            _f_due60   = df_show[(df_show["Days Until Maturity"] < -30) & (df_show["Days Until Maturity"] >= -60)]
            try:
                total_val_str = usd(df_show['Invoice Value'].sum())
            except:
                total_val_str = "$0.00"
            try:
                overdue_val_str = usd(_f_overdue['Invoice Value'].sum())
            except:
                overdue_val_str = "$0.00"
            try:
                due7_val_str = usd(_f_due7['Invoice Value'].sum())
            except:
                due7_val_str = "$0.00"
            try:
                due15_val_str = usd(_f_due15['Invoice Value'].sum())
            except:
                due15_val_str = "$0.00"
            try:
                due30_val_str = usd(_f_due30['Invoice Value'].sum())
            except:
                due30_val_str = "$0.00"
            try:
                due60_val_str = usd(_f_due60['Invoice Value'].sum())
            except:
                due60_val_str = "$0.00"

            summary_data = {
                'total': len(df_show),
                'total_val': total_val_str,
                'overdue': len(_f_overdue),
                'overdue_val': overdue_val_str,
                'due7': len(_f_due7),
                'due7_val': due7_val_str,
                'due15': len(_f_due15),
                'due15_val': due15_val_str,
                'due30': len(_f_due30),
                'due30_val': due30_val_str,
                'due60': len(_f_due60),
                'due60_val': due60_val_str,
            }

            # Build the raw df_show for PDF (before formatting)
            pdf_cols = ["Firm Name","LC No","Bank Refno","Party Name","Our Bank","Bank Name","Invoice Value",
                        "Maturity Date","Days Until Maturity","Due Status","Sales Person","Payment. Rcv Dt"]
            pdf_cols = [c for c in pdf_cols if c in df_show.columns]
            pdf_df_raw = df_show[pdf_cols].copy()
            pdf_df_raw["Maturity Date"] = pdf_df_raw["Maturity Date"].dt.strftime("%d %b %Y")
            if "Payment. Rcv Dt" in pdf_df_raw.columns:
                pdf_df_raw["Payment. Rcv Dt"] = pd.to_datetime(pdf_df_raw["Payment. Rcv Dt"], errors="coerce").dt.strftime("%d %b %Y").fillna("")
            pdf_df_raw["Days Until Maturity"] = pdf_df_raw["Days Until Maturity"].apply(lambda x: f"{int(x)}")
            inv_col_raw = df_show["Invoice Value"].copy()
            pdf_df_raw["Invoice Value"] = inv_col_raw.apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "$0.00")
            _emoji_map = {"\U0001f534 ": "", "\U0001f7e1 ": "", "\U0001f7e0 ": "", "\U0001f535 ": "", "\U0001f7e2 ": "", "\u2705 ": ""}
            if "Due Status" in pdf_df_raw.columns:
                pdf_df_raw["Due Status"] = pdf_df_raw["Due Status"].astype(str)
                for _em, _rp in _emoji_map.items():
                    pdf_df_raw["Due Status"] = pdf_df_raw["Due Status"].str.replace(_em, _rp, regex=False)
            pdf_df_raw = pdf_df_raw.fillna("")

            pdf_period = period_filter
            if isinstance(date_range, tuple) and len(date_range) == 2:
                pdf_period = f"{period_filter} ({date_range[0].strftime('%d %b %Y')} – {date_range[1].strftime('%d %b %Y')})"

            try:
                pdf_bytes_due = due_date_report_pdf(
                    pdf_df_raw, pdf_period,
                    datetime.now().strftime("%d %b %Y %H:%M:%S"),
                    summary_data)
                col_pdf_due.download_button(
                    "📄 Download Due Date Report (PDF)",
                    pdf_bytes_due,
                    file_name="due_date_tracker_report.pdf",
                    mime="application/pdf")
            except Exception as e:
                col_pdf_due.error(f"PDF error: {e}")
        else:
            col_pdf_due.warning("Install reportlab: pip install reportlab")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 10 — LEADERBOARD 🏆
# ══════════════════════════════════════════════════════════════════════════════
with t_leaderboard:
    st.markdown("## 🏆 Sales Person Leaderboard")
    st.caption("Gamified ranking with targets, trends, and performance scores")
    st.markdown("---")

    # Build leaderboard data
    if len(spg) == 0:
        st.info("No sales person data available.")
    else:
        lb = spg.copy().reset_index(drop=True)
        lb["Rank"] = lb.index + 1

        # Medal
        def medal(rank):
            if rank == 1: return "🥇"
            elif rank == 2: return "🥈"
            elif rank == 3: return "🥉"
            else: return f"#{rank}"

        lb["Medal"] = lb["Rank"].apply(medal)

        # Target = avg of all * 1.1 (aspirational target)
        avg_inv = lb["Inv"].mean()
        target  = avg_inv * 1.1
        lb["Target"] = target
        lb["Achievement %"] = (lb["Inv"] / target * 100).round(1).clip(upper=200)
        lb["Stars"] = lb["Achievement %"].apply(
            lambda p: "⭐⭐⭐⭐⭐" if p >= 150 else
                      "⭐⭐⭐⭐" if p >= 100 else
                      "⭐⭐⭐" if p >= 75 else
                      "⭐⭐" if p >= 50 else "⭐"
        )

        # Month-over-month comparison (if we have >= 2 months)
        lb["MoM"] = ""
        if len(monthly) >= 2:
            last_m = str(monthly["MonthSort"].iloc[-1])
            prev_m = str(monthly["MonthSort"].iloc[-2])
            last_sp = (df[df["MonthSort"].astype(str) == last_m][df["Sales Person"].notna()]
                       .groupby("Sales Person")["Invoice Value"].sum())
            prev_sp = (df[df["MonthSort"].astype(str) == prev_m][df["Sales Person"].notna()]
                       .groupby("Sales Person")["Invoice Value"].sum())
            for i, row in lb.iterrows():
                sp = row["Sales Person"]
                l_val = last_sp.get(sp, 0)
                p_val = prev_sp.get(sp, 0)
                if p_val > 0:
                    chg = (l_val - p_val) / p_val * 100
                    lb.at[i, "MoM"] = f"{'↑' if chg >= 0 else '↓'} {abs(chg):.0f}%"
                elif l_val > 0:
                    lb.at[i, "MoM"] = "↑ New"

        # Top 3 featured
        top3 = lb.head(3)
        p1, p2, p3 = st.columns(3)
        for col, (_, row) in zip([p2, p1, p3], [(0, top3.iloc[0]), (1, top3.iloc[1] if len(top3)>1 else top3.iloc[0]), (2, top3.iloc[2] if len(top3)>2 else top3.iloc[0])]):
            # p2=1st, p1=2nd(left), p3=3rd(right) podium style
            pass

        # Podium style for top 3
        if len(top3) >= 1:
            pm_cols = st.columns(3)
            podium_order = [1, 0, 2] if len(top3) >= 3 else list(range(len(top3)))
            heights = ["180px", "220px", "160px"]
            for i, pi in enumerate(podium_order[:len(top3)]):
                row = top3.iloc[pi]
                h   = heights[i]
                with pm_cols[i]:
                    st.markdown(f"""
                    <div class='leaderboard-card' style='text-align:center; min-height:{h};
                         background: linear-gradient(135deg, rgba(0,201,167,0.1), rgba(26,143,255,0.08));'>
                        <div style='font-size:40px; margin-bottom:4px;'>{row['Medal']}</div>
                        <div style='font-size:15px; font-weight:700; color:#e2eaf3;'>{row['Sales Person']}</div>
                        <div style='font-size:22px; font-weight:800; color:#00c9a7; margin:6px 0;'>{usd(row['Inv'])}</div>
                        <div style='font-size:12px; color:#8899aa;'>{int(row['N'])} submissions</div>
                        <div style='font-size:14px; margin-top:6px;'>{row['Stars']}</div>
                        <div style='font-size:12px; color:#ff9500; margin-top:4px;'>{row['MoM']}</div>
                        <div style='background:rgba(0,201,167,0.15); border-radius:20px; height:8px; margin:10px 0; overflow:hidden;'>
                            <div style='background:linear-gradient(90deg,#00c9a7,#1a8fff);
                                        height:100%; width:{min(100, row["Achievement %"]):.0f}%;
                                        border-radius:20px;'></div>
                        </div>
                        <div style='font-size:11px; color:#445566;'>Target: {row["Achievement %"]:.0f}% achieved</div>
                    </div>
                    """, unsafe_allow_html=True)

        st.markdown("---")

        # Full leaderboard table
        sh("📋 Full Leaderboard")
        lb_display = lb[["Medal","Sales Person","Inv","N","Paid","Pct","Stars","Achievement %","MoM"]].copy()
        lb_display["Inv"]  = lb_display["Inv"].map(lambda x: f"${x:,.2f}")
        lb_display["Pct"]  = lb_display["Pct"].map(lambda x: f"{x:.1f}%")
        lb_display["Paid"] = lb_display["Paid"].astype(int)
        lb_display["Achievement %"] = lb_display["Achievement %"].map(lambda x: f"{x:.0f}%")
        lb_display.columns = ["Rank","Sales Person","Invoice Value","Submissions","Paid","Pay Rate","Stars","Target%","MoM"]
        st.dataframe(lb_display, use_container_width=True, hide_index=True, height=400)

        st.markdown("---")

        # Bar race chart
        sh("📊 Invoice Value Comparison")
        fig_lb = px.bar(lb.head(15), x="Inv", y="Sales Person", orientation="h",
                        color="Achievement %",
                        color_continuous_scale=["#ff3b30","#ff9500","#00c9a7"],
                        text=lb.head(15)["Inv"].map(usd))
        fig_lb.update_traces(textposition="outside", textfont_size=10)
        fig_lb.update_layout(**PL_GENERAL,
            xaxis=dict(title="Invoice Value (USD)", tickformat="$.2s", gridcolor="#1a2a3a"),
            yaxis=dict(title="", autorange="reversed"),
            coloraxis_showscale=False, showlegend=False, height=450)
        st.plotly_chart(fig_lb, use_container_width=True)

        st.caption(f"🎯 Target per person: {usd(target)} (avg × 1.1) | Period: {period}")

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#445566;font-size:11px;letter-spacing:.1em;'>"
    f"Asm@2026  BANK SUBMIT HISTORY DASHBOARD &nbsp;·&nbsp; {N:,} RECORDS &nbsp;·&nbsp; {period} "
    f"&nbsp;·&nbsp; 🏦 Smart Dashboard v2.0</p>", unsafe_allow_html=True)
