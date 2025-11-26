import os
import uuid
import io
import threading
import zipfile
import time
import random
import string
import secrets
from datetime import datetime, timedelta, timezone, date
from functools import wraps
from threading import Lock
import mimetypes

import openpyxl
from openpyxl import Workbook

from flask import (
    Flask, request, jsonify, send_file, render_template
)
from flask_sqlalchemy import SQLAlchemy
from flask_jwt_extended import (
    JWTManager, create_access_token, jwt_required,
    get_jwt_identity, verify_jwt_in_request, get_jwt
)
from flask_socketio import SocketIO
from flask_cors import CORS
from sqlalchemy import func, and_, or_
from sqlalchemy.exc import IntegrityError
from werkzeug.utils import secure_filename
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import smtplib
from email.message import EmailMessage

# Optional libraries
try:
    import stripe
    STRIPE_AVAILABLE = True
except Exception:
    stripe = None
    STRIPE_AVAILABLE = False

try:
    import bcrypt
    BCRYPT_AVAILABLE = True
except Exception:
    bcrypt = None
    BCRYPT_AVAILABLE = False

try:
    import razorpay
    RAZORPAY_AVAILABLE = True
except Exception:
    razorpay = None
    RAZORPAY_AVAILABLE = False

try:
    from twilio.rest import Client as TwilioClient
    TWILIO_AVAILABLE = True
except Exception:
    TwilioClient = None
    TWILIO_AVAILABLE = False

# ----------- CONFIG -----------

from dotenv import load_dotenv
load_dotenv()

app = Flask(__name__)

app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv(
    "DATABASE_URI",
    "mysql+pymysql://root:your_sql_password.@localhost:3306/oralcare_db"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["JWT_SECRET_KEY"] = os.getenv("JWT_SECRET_KEY", "supersecret")
app.config["JWT_ACCESS_TOKEN_EXPIRES"] = int(os.getenv("JWT_EXP_SECONDS", 60 * 60 * 24))

app.config["UPLOAD_FOLDER"] = os.getenv("UPLOAD_FOLDER", "./uploads")
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

app.config["PROFILE_UPLOADS_DIR"] = os.path.join(app.config["UPLOAD_FOLDER"], "profiles")
app.config["CERT_UPLOADS_DIR"] = os.path.join(app.config["UPLOAD_FOLDER"], "certs")
os.makedirs(app.config["PROFILE_UPLOADS_DIR"], exist_ok=True)
os.makedirs(app.config["CERT_UPLOADS_DIR"], exist_ok=True)

app.config["REPORTS_DIR"] = os.path.join(os.getcwd(), "reports")
os.makedirs(app.config["REPORTS_DIR"], exist_ok=True)

# For prescriptions and clinic PDFs
app.config["PRESCRIPTIONS_DIR"] = os.path.join(os.getcwd(), "prescriptions")
os.makedirs(app.config["PRESCRIPTIONS_DIR"], exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.config["ALLOWED_IMAGE_EXTENSIONS"] = {"png", "jpg", "jpeg", "bmp", "gif"}
app.config["ALLOWED_CERT_EXTENSIONS"] = {"pdf", "png", "jpg", "jpeg"}

FRONTEND_BASE_URL = os.getenv("FRONTEND_BASE_URL", "http://localhost:3000")
STRIPE_SECRET_KEY = os.getenv("STRIPE_SECRET_KEY", "")
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "")
if STRIPE_AVAILABLE and STRIPE_SECRET_KEY:
    stripe.api_key = STRIPE_SECRET_KEY

RAZORPAY_KEY_ID = os.getenv("RAZORPAY_KEY_ID", "")
RAZORPAY_KEY_SECRET = os.getenv("RAZORPAY_KEY_SECRET", "")
if RAZORPAY_AVAILABLE and RAZORPAY_KEY_ID and RAZORPAY_KEY_SECRET:
    razorpay_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET))
else:
    razorpay_client = None

# Email settings (for smtplib)
EMAIL_HOST = os.getenv("EMAIL_HOST", "smtp.gmail.com")
EMAIL_PORT = int(os.getenv("EMAIL_PORT", "587"))
EMAIL_USER = os.getenv("EMAIL_USER", "oralcare.demo@gmail.com")
EMAIL_PASS = os.getenv("EMAIL_PASS", "your-app-password")
EMAIL_USE_TLS = os.getenv("EMAIL_USE_TLS", "1") == "1"

# Twilio (optional)
TWILIO_SID = os.getenv("TWILIO_SID", "")
TWILIO_AUTH = os.getenv("TWILIO_AUTH", "")
TWILIO_FROM = os.getenv("TWILIO_FROM", "")

db = SQLAlchemy(app)
jwt = JWTManager(app)
socketio = SocketIO(app, cors_allowed_origins="*")
CORS(app)

queue_lock = Lock()
OTP_VALIDITY_SECONDS = 5 * 60  # 5 mins

# IST timezone for clinic
IST = timezone(timedelta(hours=5, minutes=30))

# ----------- TRANSLATION SETUP -----------

TRANSLATIONS = {
    "en": {
        "otp_sent": "OTP sent to your contact.",
        "otp_verified": "OTP verified. Login successful.",
        "invalid_otp": "Invalid or expired OTP.",
        "account_created": "Account created and logged in.",
        "already_exists": "User already exists.",
        "user_not_found": "User not found.",
        "login_success": "Login successful.",
        "invalid_credentials": "Invalid credentials.",
        "profile_updated": "Profile updated.",
        "upload_success": "File uploaded.",
        "not_allowed": "Forbidden.",
        "operation_success": "Operation successful.",
        "user_deleted": "User deleted.",
    },
    "es": {
        "otp_sent": "OTP enviado a su contacto.",
        "otp_verified": "OTP verificado. Inicio de sesión exitoso.",
        "invalid_otp": "OTP inválido o caducado.",
        "account_created": "Cuenta creada e inició sesión.",
        "already_exists": "El usuario ya existe.",
        "user_not_found": "Usuario no encontrado.",
        "login_success": "Inicio de sesión exitoso.",
        "invalid_credentials": "Credenciales inválidas.",
        "profile_updated": "Perfil actualizado.",
        "upload_success": "Archivo subido.",
        "not_allowed": "Prohibido.",
        "operation_success": "Operación exitosa.",
        "user_deleted": "Usuario eliminado.",
    },
}


def tr(key, lang):
    return TRANSLATIONS.get(lang, TRANSLATIONS["en"]).get(key, key)


def parse_lang():
    lang = request.args.get("lang", "en")
    if lang not in TRANSLATIONS:
        lang = "en"
    return lang

# ----------- MODELS -----------

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(150), nullable=False)
    email = db.Column(db.String(150), unique=True, nullable=False)
    phone = db.Column(db.String(50), unique=True, nullable=True)
    password = db.Column(db.Text, nullable=False)
    role = db.Column(db.String(50), nullable=False, default="patient")
    dob = db.Column(db.Date, nullable=True)
    address = db.Column(db.Text)
    profile_pic = db.Column(db.String(256))
    certificate_file = db.Column(db.String(256))


class Service(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(150))
    price = db.Column(db.Integer)


class Appointment(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    patient_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    doctor_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    service_id = db.Column(db.Integer, db.ForeignKey("service.id"))
    requested_time = db.Column(db.DateTime(timezone=True))
    status = db.Column(db.String(20), default="requested")  # requested, accepted, completed, cancelled, rescheduled
    meeting_link = db.Column(db.String(300))
    payment_status = db.Column(db.String(20), default="pending")  # pending, paid, failed
    queue_number = db.Column(db.Integer)  # global queue number for the day/clinic

    # Single-doctor clinic enhancements
    visit_type = db.Column(db.String(20), default="online")  # online / walkin
    queue_status = db.Column(db.String(20), default="waiting")  # waiting, in_progress, done, skipped

    doctor = db.relationship("User", foreign_keys=[doctor_id])
    patient = db.relationship("User", foreign_keys=[patient_id])


class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    appointment_id = db.Column(db.Integer, db.ForeignKey("appointment.id"), nullable=False)
    diagnosis = db.Column(db.String(255))
    prescription = db.Column(db.String(255))
    notes = db.Column(db.Text)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    filename = db.Column(db.String(255))
    file_data = db.Column(db.LargeBinary)
    file_path = db.Column(db.String(512), nullable=False)
    appointment = db.relationship("Appointment", backref=db.backref("reports", lazy=True))


class OTP(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    contact = db.Column(db.String(150), nullable=False, index=True)
    code = db.Column(db.String(12), nullable=False)
    mode = db.Column(db.String(10), nullable=False)  # "email" or "sms"
    expires_at = db.Column(db.DateTime(timezone=True))
    used = db.Column(db.Boolean, default=False)
    auto_create = db.Column(db.Boolean, default=False)


class Prescription(db.Model):
    """
    Digital prescription generated by the doctor for a specific appointment.
    """
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    appointment_id = db.Column(db.Integer, db.ForeignKey("appointment.id"), nullable=False)
    doctor_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    patient_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)

    diagnosis = db.Column(db.Text)
    notes = db.Column(db.Text)
    medicines = db.Column(db.JSON)  # [{name, dose, frequency, duration, remarks}, ...]

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    pdf_path = db.Column(db.String(512))

    appointment = db.relationship("Appointment", backref=db.backref("prescriptions", lazy=True))
    doctor = db.relationship("User", foreign_keys=[doctor_id], lazy=True)
    patient = db.relationship("User", foreign_keys=[patient_id], lazy=True)


class Income(db.Model):
    """
    Income entries – usually auto-created from paid appointments.
    """
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    appointment_id = db.Column(db.Integer, db.ForeignKey("appointment.id"), nullable=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    amount = db.Column(db.Float, nullable=False, default=0.0)
    source = db.Column(db.String(100), default="appointment")  # appointment, other
    notes = db.Column(db.Text)

    appointment = db.relationship("Appointment", backref=db.backref("income_entries", lazy=True))


class Expense(db.Model):
    """
    Manual expense entries – consumables, rent, salaries, etc.
    """
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    type = db.Column(db.String(200))
    amount = db.Column(db.Float, nullable=False, default=0.0)
    notes = db.Column(db.Text)


class InventoryItem(db.Model):
    """
    Simple clinic inventory for low-stock & expiry warnings.
    """
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(200), nullable=False)
    quantity = db.Column(db.Integer, nullable=False, default=0)
    min_required = db.Column(db.Integer, nullable=False, default=0)
    expiry_date = db.Column(db.Date, nullable=True)
    unit = db.Column(db.String(50), default="unit")  # e.g. box, pack, pcs


# ----------- HELPERS -----------

def _to_int(v):
    try:
        return int(v)
    except Exception:
        return None


def hash_password(password):
    if not BCRYPT_AVAILABLE:
        return password
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()


def check_password(password, hashed):
    if BCRYPT_AVAILABLE:
        try:
            return bcrypt.checkpw(password.encode(), hashed.encode())
        except Exception:
            pass
    return password == hashed


def create_jwt(user):
    return create_access_token(identity=str(user.id), additional_claims={"role": user.role})


def role_required(allowed_roles):
    def wrapper(fn):
        @wraps(fn)
        def decorator(*args, **kwargs):
            verify_jwt_in_request()
            claims = get_jwt()
            user_role = claims.get("role")
            if user_role not in allowed_roles:
                return jsonify({"msg": tr("not_allowed", parse_lang())}), 403
            return fn(*args, **kwargs)
        return decorator
    return wrapper


def allowed_file(filename, mode="img"):
    ALLOWED = app.config["ALLOWED_IMAGE_EXTENSIONS"] if mode == "img" else app.config["ALLOWED_CERT_EXTENSIONS"]
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED


def safe_text_for_pdf(s):
    if s is None:
        return ""
    return str(s).encode("latin-1", "replace").decode("latin-1")


def gen_otp_code():
    return "".join(secrets.choice(string.digits) for _ in range(6))


def send_email(to, subject, body):
    try:
        em = EmailMessage()
        em["From"] = EMAIL_USER
        em["To"] = to
        em["Subject"] = subject
        em.set_content(body)
        with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as smtp:
            if EMAIL_USE_TLS:
                smtp.starttls()
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(em)
        return True
    except Exception as e:
        print("Email sending failed:", e)
        return False


def send_sms(to, msg):
    if not (TWILIO_AVAILABLE and TWILIO_SID and TWILIO_AUTH and TWILIO_FROM):
        print("Twilio not set up or not available.")
        return False
    try:
        client = TwilioClient(TWILIO_SID, TWILIO_AUTH)
        client.messages.create(
            body=msg, from_=TWILIO_FROM, to=to
        )
        return True
    except Exception as e:
        print("Twilio SMS failed:", e)
        return False


def notify_user(user, subject, body, sms_msg=None):
    sent = False
    if user.email:
        sent = send_email(user.email, subject, body)
    if sms_msg and user.phone:
        send_sms(user.phone, sms_msg)
    return sent


def pag_sort_query(query, model, allowed_sort, default_sort="id"):
    page = _to_int(request.args.get("page", 1)) or 1
    per_page = min(max(_to_int(request.args.get("per_page", 20)) or 20, 1), 100)
    sort = request.args.get("sort", default_sort)
    if sort.startswith("-"):
        col = getattr(model, sort[1:], None)
        if col is not None:
            query = query.order_by(col.desc())
    else:
        col = getattr(model, sort, None)
        if col is not None:
            query = query.order_by(col.asc())
    return query.offset((page - 1) * per_page).limit(per_page)


def parse_client_datetime(requested_time_raw):
    """
    Parse datetime sent from frontend and return:
      - dt_ist: time in IST
      - dt_utc: same moment in UTC (for DB storage)
    Accepts:
      - "YYYY-MM-DDTHH:MM"
      - full ISO, with or without milliseconds, with optional 'Z' or offset
    """
    if not isinstance(requested_time_raw, str):
        raise ValueError("requested_time must be string")

    s = requested_time_raw.strip()

    # normalize 'Z' to +00:00
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"

    dt = datetime.fromisoformat(s)

    if dt.tzinfo is None:
        # Treat naive time as IST (clinic local time)
        dt_ist = dt.replace(tzinfo=IST)
    else:
        # Convert whatever zone to IST
        dt_ist = dt.astimezone(IST)

    dt_utc = dt_ist.astimezone(timezone.utc)
    return dt_ist, dt_utc


# ----------- SEED ------------

with app.app_context():
    db.create_all()

    # Ensure single default doctor exists
    if db.session.query(User).filter_by(role="doctor").first() is None:
        doctor = User(
            name="doctor",
            email="doctor@gmail.com",
            password=hash_password("123"),
            role="doctor"
        )
        db.session.add(doctor)
        db.session.commit()

    # Seed basic services if none
    if db.session.query(Service).count() == 0:
        services = [
            "Orthodontics", "Root Canal Treatment", "Pediatric Dentistry", "Periodontal Care",
            "Laser Treatment", "Dental Implants", "Oral Surgery", "Teeth Whitening",
            "Dental Crowns", "Dentures", "Low Radiation X-Rays", "Cosmetic Dentistry",
            "Smile Designing", "Crowns & Bridges", "Wisdom Tooth Removal",
            "Full Mouth Rehabilitation", "Prosthodontics", "Maxillofacial Surgery",
            "Aligners", "Dental Jewellery"
        ]
        db.session.add_all([Service(name=s, price=1000 + i * 200) for i, s in enumerate(services)])
        db.session.commit()

    # Optional: Seed some inventory for demo
    if db.session.query(InventoryItem).count() == 0:
        demo_items = [
            InventoryItem(name="Gloves", quantity=200, min_required=100, unit="pair"),
            InventoryItem(name="Masks", quantity=150, min_required=80, unit="pcs"),
            InventoryItem(name="Composite Filling Material", quantity=20, min_required=10, unit="syringe"),
        ]
        db.session.add_all(demo_items)
        db.session.commit()
    # Ensure default admin exists
    if db.session.query(User).filter_by(role="admin").first() is None:
        admin = User(
            name="Admin",
            email="admin@gmail.com",
            password=hash_password("admin123"),
            role="admin"
        )
        db.session.add(admin)
        db.session.commit()
        print("✔ Default admin created → admin@gmail.com / admin123")


# ----------- AUTH & USER -----------

@app.route("/register", methods=["POST"])
def register():
    data = request.json or {}
    lang = parse_lang()
    if "email" not in data or "password" not in data:
        return jsonify({"msg": "Email and password are required"}), 400

    name = data.get("name") or "Unnamed"
    email = data["email"].strip().lower()
    password = data["password"]
    phone = data.get("phone")
    role = data.get("role", "patient")

    if db.session.query(User).filter(or_(User.email == email, User.phone == phone)).first():
        return jsonify({"msg": tr("already_exists", lang)}), 400

    # Single doctor clinic: allow only one doctor
    if role == "doctor" and db.session.query(User).filter_by(role="doctor").count() >= 1:
        return jsonify({"msg": "Only one doctor account allowed."}), 403

    user = User(name=name, email=email, phone=phone, password=hash_password(password), role=role)
    db.session.add(user)
    db.session.commit()
    return jsonify({"msg": tr("account_created", lang), "id": user.id}), 201


@app.route("/login", methods=["POST"])
def login():
    data = request.json or {}
    lang = parse_lang()
    email = data.get("email", "").strip().lower()
    password = data.get("password") or ""
    user = db.session.query(User).filter_by(email=email).first()
    if not user or not check_password(password, user.password):
        return jsonify({"msg": tr("invalid_credentials", lang)}), 401
    token = create_jwt(user)
    return jsonify({"token": token, "role": user.role, "name": user.name, "id": user.id})


# ----------- OTP LOGIN FLOW ------------

@app.route("/auth/send-otp", methods=["POST"])
def send_otp():
    data = request.json or {}
    lang = parse_lang()
    mode = data.get("mode", "email")
    contact = (data.get("email") or data.get("phone") or "").lower()
    auto_create = bool(data.get("auto_create", False))

    if mode == "sms" and not contact:
        return jsonify({"msg": "phone required for sms"}), 400

    user = db.session.query(User).filter(
        User.email == contact if mode == "email" else User.phone == contact
    ).first()

    if not user and not auto_create:
        return jsonify({"msg": tr("user_not_found", lang)}), 404

    code = gen_otp_code()
    ot = OTP(
        contact=contact,
        code=code,
        mode=mode,
        expires_at=(datetime.utcnow().replace(tzinfo=timezone.utc) + timedelta(seconds=OTP_VALIDITY_SECONDS)),
        auto_create=auto_create
    )
    db.session.add(ot)
    db.session.commit()

    body = f"OralCare OTP: {code}\nThis OTP will expire in 5 minutes."
    if mode == "email":
        send_email(contact, "OralCare OTP", body)
    elif mode == "sms":
        send_sms(contact, f"Your OTP: {code}")

    return jsonify({"msg": tr("otp_sent", lang)})


@app.route("/auth/verify-otp", methods=["POST"])
def verify_otp():
    data = request.json or {}
    lang = parse_lang()
    mode = data.get("mode", "email")
    contact = (data.get("email") or data.get("phone") or "").lower()
    code = data.get("otp")
    auto_create = bool(data.get("auto_create", False))

    otp_row = db.session.query(OTP).filter(
        OTP.contact == contact,
        OTP.code == code,
        OTP.mode == mode,
        OTP.used == False
    ).order_by(OTP.id.desc()).first()

    # --- FIXED TIMEZONE SAFE CHECK ---
    now_utc = datetime.now(timezone.utc)

    if not otp_row:
        return jsonify({"msg": tr("invalid_otp", lang)}), 400

    expires = otp_row.expires_at
    if expires.tzinfo is None:
        expires = expires.replace(tzinfo=timezone.utc)

    if expires < now_utc:
        return jsonify({"msg": tr("invalid_otp", lang)}), 400

    # ---------------------------------

    user = db.session.query(User).filter(
        User.email == contact if mode == "email" else User.phone == contact
    ).first()

    if not user and auto_create:
        user = User(
            name="User",
            email=contact if mode == "email" else None,
            phone=contact if mode == "sms" else None,
            password=hash_password(gen_otp_code()),
            role="patient"
        )
        db.session.add(user)

    if not user:
        return jsonify({"msg": tr("user_not_found", lang)}), 404

    otp_row.used = True
    db.session.commit()

    token = create_jwt(user)
    return jsonify({
        "token": token,
        "id": user.id,
        "role": user.role,
        "name": user.name,
        "msg": tr("otp_verified", lang)
    })



# ----------- PRESCRIPTION PDF HELPER -----------

def generate_prescription_pdf(presc: Prescription):
    """
    Generate a PDF for the prescription and update its pdf_path.
    """
    try:
        appt = db.session.get(Appointment, presc.appointment_id)
        doctor = db.session.get(User, presc.doctor_id)
        patient = db.session.get(User, presc.patient_id)
        service = db.session.get(Service, appt.service_id) if appt and appt.service_id else None

        if not appt or not doctor or not patient:
            return

        filename = f"prescription_{presc.id}.pdf"
        filepath = os.path.join(app.config["PRESCRIPTIONS_DIR"], filename)

        buf = canvas.Canvas(filepath, pagesize=letter)
        y = 760

        # Header
        buf.setFont("Helvetica-Bold", 16)
        buf.drawString(60, y, safe_text_for_pdf("OralCare Clinic - Prescription"))
        y -= 30

        buf.setFont("Helvetica", 10)
        buf.drawString(60, y, safe_text_for_pdf(f"Doctor: {doctor.name}"))
        y -= 15
        buf.drawString(60, y, safe_text_for_pdf(f"Patient: {patient.name}"))
        y -= 15
        if service:
            buf.drawString(60, y, safe_text_for_pdf(f"Service: {service.name}"))
            y -= 15
        if appt.requested_time:
            buf.drawString(60, y, safe_text_for_pdf(f"Date: {appt.requested_time.date().isoformat()}"))
            y -= 25

        # Diagnosis & notes
        if presc.diagnosis:
            buf.setFont("Helvetica-Bold", 11)
            buf.drawString(60, y, "Diagnosis:")
            y -= 15
            buf.setFont("Helvetica", 10)
            for line in str(presc.diagnosis).splitlines():
                buf.drawString(80, y, safe_text_for_pdf(line))
                y -= 12
            y -= 10

        if presc.notes:
            buf.setFont("Helvetica-Bold", 11)
            buf.drawString(60, y, "Notes:")
            y -= 15
            buf.setFont("Helvetica", 10)
            for line in str(presc.notes).splitlines():
                buf.drawString(80, y, safe_text_for_pdf(line))
                y -= 12
            y -= 10

        # Medicines
        meds = presc.medicines or []
        if meds:
            buf.setFont("Helvetica-Bold", 11)
            buf.drawString(60, y, "Medicines:")
            y -= 18
            buf.setFont("Helvetica", 10)

            for idx, m in enumerate(meds, start=1):
                if y < 80:
                    buf.showPage()
                    y = 760
                    buf.setFont("Helvetica", 10)

                name = m.get("name", "")
                dose = m.get("dose", "")
                freq = m.get("frequency", "")
                dur = m.get("duration", "")
                remark = m.get("remarks", "")

                line = f"{idx}. {name}  {dose}  {freq}  for {dur}"
                buf.drawString(80, y, safe_text_for_pdf(line))
                y -= 12
                if remark:
                    buf.drawString(100, y, safe_text_for_pdf(f"Note: {remark}"))
                    y -= 12
                y -= 4

        # Footer
        if y < 120:
            buf.showPage()
            y = 760

        buf.setFont("Helvetica", 9)
        buf.drawString(60, 80, safe_text_for_pdf("This is a computer-generated prescription."))
        buf.drawString(60, 65, safe_text_for_pdf("Please follow up with your dentist as advised."))

        buf.showPage()
        buf.save()

        presc.pdf_path = filepath
        db.session.commit()
    except Exception as e:
        print("Failed to generate prescription PDF:", e)


def ensure_income_for_appointment(appt: Appointment):
    """Ensure Income entry exists when appointment is paid."""
    try:
        if not appt:
            return
        existing = db.session.query(Income).filter_by(appointment_id=appt.id).first()
        if existing:
            return
        service = db.session.get(Service, appt.service_id) if appt.service_id else None
        amount = float(service.price) if service else 0.0
        inc = Income(
            appointment_id=appt.id,
            amount=amount,
            date=(appt.requested_time.date() if appt.requested_time else date.today()),
            source="appointment",
            notes=f"Payment for appointment #{appt.id}"
        )
        db.session.add(inc)
        db.session.commit()
    except Exception as e:
        print("ensure_income_for_appointment error:", e)


# ----------- ADMIN DASHBOARD & ADMIN APIs ------------

@app.route("/dashboard/admin", methods=["GET"])
@role_required(["admin"])
def admin_dashboard():
    total_users = db.session.query(User).count()
    total_doctors = db.session.query(User).filter_by(role="doctor").count()
    total_patients = db.session.query(User).filter_by(role="patient").count()
    total_appts = db.session.query(Appointment).count()

    revenue = db.session.query(func.sum(Income.amount)).scalar() or 0.0

    one_week_ago = datetime.utcnow() - timedelta(days=7)
    weekly_counts = db.session.query(
        func.date(Appointment.requested_time),
        func.count(Appointment.id)
    ).filter(
        Appointment.requested_time >= one_week_ago
    ).group_by(
        func.date(Appointment.requested_time)
    ).all()
    week_stats = [{"date": str(d), "count": c} for d, c in weekly_counts]

    new_regs = db.session.query(User).filter(User.role == "patient").order_by(User.id.desc()).limit(5).all()
    recent = [{"id": u.id, "name": u.name, "email": u.email, "date": u.id} for u in new_regs]

    return jsonify({
        "total_users": total_users,
        "total_doctors": total_doctors,
        "total_patients": total_patients,
        "total_appointments": total_appts,
        "revenue": revenue,
        "weekly_appointment_stats": week_stats,
        "recent_registrations": recent
    })


@app.route("/admin/users", methods=["GET"])
@role_required(["admin"])
def list_users():
    q = db.session.query(User)
    q = pag_sort_query(q, User, ["id", "email", "name"], "id")
    users = q.all()
    return jsonify([{
        "id": u.id,
        "name": u.name,
        "email": u.email,
        "role": u.role,
        "phone": u.phone,
        "dob": u.dob.isoformat() if u.dob else None
    } for u in users])


@app.route("/admin/create-doctor", methods=["POST"])
@role_required(["admin"])
def create_doctor():
    data = request.json or {}
    email, name, password = data.get("email"), data.get("name"), data.get("password")
    if not (email and name and password):
        return jsonify({"msg": "All required fields missing"}), 400
    if db.session.query(User).filter_by(email=email).first():
        return jsonify({"msg": "Email exists"}), 409
    if db.session.query(User).filter_by(role="doctor").count() >= 1:
        return jsonify({"msg": "Only one doctor allowed"}), 400
    doc = User(email=email, name=name, password=hash_password(password), role="doctor")
    db.session.add(doc)
    db.session.commit()
    return jsonify({"msg": "Doctor created", "id": doc.id}), 201


@app.route("/admin/delete-user/<id>", methods=["DELETE"])
@role_required(["admin"])
def delete_user(id):
    user = db.session.get(User, _to_int(id))
    if not user:
        return jsonify({"msg": "User does not exist"}), 404
    db.session.delete(user)
    db.session.commit()
    return jsonify({"msg": tr("user_deleted", parse_lang())})


# ----------- APPOINTMENTS (ONLINE BOOKING + WALK-IN) -----------

@app.route("/appointments/book", methods=["POST"])
@role_required(["patient"])
def book_appointment():
    """
    Option A:
      - Single doctor auto-selected
      - Patient sends: service_id, requested_time
      - requested_time can be:
          "YYYY-MM-DDTHH:MM"   (local IST)
          or full ISO string (frontend using toISOString)
      - Booking allowed only between 09:00–18:00 IST
      - Prevents double booking for same doctor & time
    """
    data = request.json or {}
    patient_id = int(get_jwt_identity())
    service_id = data.get("service_id")
    requested_time_raw = data.get("requested_time")

    if not service_id or not requested_time_raw:
        return jsonify({"msg": "service_id and requested_time required"}), 400

    # Auto-select the only doctor
    doctor = db.session.query(User).filter_by(role="doctor").first()
    if not doctor:
        return jsonify({"msg": "No doctor configured"}), 500
    doctor_id = doctor.id

    # Parse datetime from client
    try:
        dt_ist, dt_utc = parse_client_datetime(requested_time_raw)
    except Exception:
        return jsonify({"msg": "Invalid datetime format"}), 400

    # Booking allowed only from 09:00 to <18:00 IST
    hour = dt_ist.hour
    if hour < 9 or hour >= 18:
        return jsonify({"msg": "Booking allowed only between 09:00 and 18:00"}), 400

    # Prevent double booking (same doctor, same exact time)
    exists = db.session.query(Appointment).filter(
        Appointment.doctor_id == doctor_id,
        Appointment.requested_time == dt_utc,
        Appointment.status != "cancelled"
    ).first()

    if exists:
        return jsonify({"msg": "Time already booked"}), 400

    # Create appointment with next queue number
    with queue_lock:
        max_q = db.session.query(func.max(Appointment.queue_number)).scalar()
        next_queue = (max_q or 0) + 1

        appt = Appointment(
            patient_id=patient_id,
            doctor_id=doctor_id,
            service_id=int(service_id),
            requested_time=dt_utc,
            queue_number=next_queue,
            status="requested",
            visit_type="online",
            queue_status="waiting"
        )
        db.session.add(appt)
        db.session.commit()

    socketio.emit("new_appointment", {
        "id": appt.id,
        "queue_number": next_queue,
        "service_id": appt.service_id,
        "status": appt.status,
        "visit_type": appt.visit_type
    })

    return jsonify({
        "msg": "Appointment booked successfully",
        "queue_number": next_queue,
        "id": appt.id
    }), 201


@app.route("/appointments/walkin", methods=["POST"])
def create_walkin_appointment():
    """
    In-clinic quick create for walk-in patients (no login).
    """
    data = request.json or {}
    name = data.get("name")
    phone = data.get("phone")
    service_id = data.get("service_id")

    if not (name and phone and service_id):
        return jsonify({"msg": "name, phone and service_id required"}), 400

    # Single clinic doctor
    doctor = db.session.query(User).filter_by(role="doctor").first()
    if not doctor:
        return jsonify({"msg": "No doctor configured"}), 500

    # Find or create patient by phone
    patient = db.session.query(User).filter_by(phone=phone).first()
    if not patient:
        patient = User(
            name=name,
            email=f"walkin_{phone}@example.com",
            phone=phone,
            password=hash_password(gen_otp_code()),
            role="patient"
        )
        db.session.add(patient)
        db.session.flush()

    now_utc = datetime.utcnow().replace(tzinfo=timezone.utc)
    with queue_lock:
        max_q = db.session.query(func.max(Appointment.queue_number)).scalar()
        next_queue = (max_q or 0) + 1
        appt = Appointment(
            patient_id=patient.id,
            doctor_id=doctor.id,
            service_id=int(service_id),
            requested_time=now_utc,
            queue_number=next_queue,
            status="requested",
            visit_type="walkin",
            queue_status="waiting"
        )
        db.session.add(appt)
        db.session.commit()

    socketio.emit("new_appointment", {
        "id": appt.id,
        "queue_number": next_queue,
        "doctor": doctor.name,
        "service_id": appt.service_id,
        "status": appt.status,
        "visit_type": appt.visit_type
    })

    return jsonify({
        "msg": "Walk-in appointment created",
        "appointment_id": appt.id,
        "token": next_queue
    }), 201


# ----------- SEARCH/FILTER -----------

@app.route("/search/patients")
@role_required(["admin"])
def search_patients():
    q = db.session.query(User).filter_by(role="patient")
    name = request.args.get("name")
    email = request.args.get("email")
    if name:
        q = q.filter(User.name.ilike(f"%{name}%"))
    if email:
        q = q.filter(User.email.ilike(f"%{email}%"))
    return jsonify([
        {
            "id": u.id,
            "name": u.name,
            "email": u.email,
            "phone": u.phone,
            "dob": u.dob.isoformat() if u.dob else None
        }
        for u in pag_sort_query(q, User, ["id", "name", "email"], "id").all()
    ])


@app.route("/search/doctors")
@jwt_required()
def search_doctors():
    q = db.session.query(User).filter_by(role="doctor")
    name = request.args.get("name")
    if name:
        q = q.filter(User.name.ilike(f"%{name}%"))
    return jsonify([
        {"id": u.id, "name": u.name, "email": u.email, "phone": u.phone}
        for u in pag_sort_query(q, User, ["id", "name"], "id").all()
    ])


@app.route("/search/appointments")
@role_required(["admin"])
def search_appointments():
    q = db.session.query(Appointment)
    date_str = request.args.get("date")
    doctor_id = request.args.get("doctor_id")
    status = request.args.get("status")
    if doctor_id:
        q = q.filter(Appointment.doctor_id == _to_int(doctor_id))
    if status:
        q = q.filter(Appointment.status == status)
    if date_str:
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d").date()
            q = q.filter(func.date(Appointment.requested_time) == dt)
        except Exception:
            pass
    q = pag_sort_query(q, Appointment, ["id", "requested_time", "status"], "id")
    return jsonify([
        {
            "id": a.id,
            "doctor_id": a.doctor_id,
            "patient_id": a.patient_id,
            "service_id": a.service_id,
            "requested_time": a.requested_time.isoformat() if a.requested_time else None,
            "status": a.status,
            "payment_status": a.payment_status,
            "queue_number": a.queue_number,
            "visit_type": a.visit_type,
            "queue_status": a.queue_status
        }
        for a in q.all()
    ])


# ----------- EMAIL/SMS NOTIFICATIONS -------------

def notify_event(event, appointment, extra_msg=None):
    try:
        pt = db.session.get(User, appointment.patient_id)
        doc = db.session.get(User, appointment.doctor_id) if appointment.doctor_id else None
        svc = db.session.get(Service, appointment.service_id) if appointment.service_id else None

        subj = f"OralCare Clinic: {event}"

        # Format date nicely
        appt_time = (
            appointment.requested_time.strftime("%d %b %Y, %I:%M %p")
            if appointment.requested_time else "N/A"
        )

        # --- MESSAGE BUILDER ---
        if event == "Appointment booked":
            msg = (
                f"Dear {pt.name},\n\n"
                f"Your appointment with Dr. {doc.name if doc else 'our doctor'} "
                f"for {svc.name if svc else 'consultation'} has been booked.\n\n"
                f"Date & Time: {appt_time}\n\n"
                "Thank you!"
            )

        elif event == "Appointment rescheduled":
            msg = f"Your appointment has been rescheduled to {appt_time}."

        elif event == "Appointment accepted":
            msg = (
                "Your appointment has been accepted.\n"
                f"Meeting link: {appointment.meeting_link or 'N/A'}"
            )

        elif event == "Appointment cancelled":
            msg = "Your appointment has been cancelled."

        elif event == "Payment successful":
            msg = (
                f"Payment for appointment #{appointment.id} was completed.\n"
                "Thank you!"
            )

        elif event == "Report uploaded":
            msg = "A new report has been uploaded to your account."

        else:
            msg = extra_msg or "You have a new notification."

        # --- SEND EMAIL / SMS ---
        if pt and pt.email:
            print(f"[EMAIL] Sending to {pt.email} → {event}")
            send_email(pt.email, subj, msg)

        if pt and pt.phone:
            print(f"[SMS] Sending to {pt.phone} → {event}")
            send_sms(pt.phone, msg)

    except Exception as ex:
        print(f"[notify_event] ERROR: {ex}")



# ----------- QUEUE / TOKEN MANAGEMENT ------------

@app.route("/queue/today", methods=["GET"])
@role_required(["doctor", "admin"])
def queue_today():
    today = datetime.utcnow().date()
    appts = db.session.query(Appointment).filter(
        func.date(Appointment.requested_time) == today,
        Appointment.status != "cancelled"
    ).order_by(Appointment.queue_number.asc()).all()

    return jsonify([
        {
            "id": a.id,
            "patient_id": a.patient_id,
            "queue_number": a.queue_number,
            "status": a.status,
            "queue_status": a.queue_status,
            "visit_type": a.visit_type,
            "requested_time": a.requested_time.isoformat() if a.requested_time else None
        }
        for a in appts
    ])


@app.route("/queue/next", methods=["POST"])
@role_required(["doctor"])
def queue_next():
    today = datetime.utcnow().date()

    appt = db.session.query(Appointment).filter(
        func.date(Appointment.requested_time) == today,
        Appointment.queue_status == "waiting",
        Appointment.status != "cancelled"
    ).order_by(
        db.case((Appointment.queue_number == None, 1), else_=0),
        Appointment.queue_number.asc()
    ).first()

    if not appt:
        return jsonify({"msg": "No waiting patients in queue"}), 404

    appt.queue_status = "in_progress"
    if appt.status == "requested":
        appt.status = "accepted"
    db.session.commit()

    socketio.emit("queue_update", {
        "event": "next",
        "appointment_id": appt.id,
        "queue_number": appt.queue_number
    })
    return jsonify({
        "msg": "Next patient set to in_progress",
        "appointment_id": appt.id,
        "queue_number": appt.queue_number
    })
@app.route("/queue/complete-and-next", methods=["POST"])
@role_required(["doctor"])
def queue_complete_and_next():
    today = datetime.utcnow().date()

    # 1. Find the in-progress appointment
    current = db.session.query(Appointment).filter(
        func.date(Appointment.requested_time) == today,
        Appointment.queue_status == "in_progress",
        Appointment.status != "cancelled"
    ).first()

    if not current:
        return jsonify({"msg": "No in-progress appointment"}), 404

    # Mark current appointment completed
    current.status = "completed"
    current.queue_status = "done"
    db.session.commit()

    socketio.emit("appointment_update", {
        "id": current.id,
        "status": "completed"
    })

    # 2. Find next patient in waiting list
    next_appt = db.session.query(Appointment).filter(
        func.date(Appointment.requested_time) == today,
        Appointment.queue_status == "waiting",
        Appointment.status != "cancelled"
    ).order_by(Appointment.queue_number.asc()).first()

    if not next_appt:
        return jsonify({
            "msg": "Completed current. No more patients in queue.",
            "completed": {
                "appointment_id": current.id,
                "queue_number": current.queue_number
            },
            "next": None
        })

    # Move next appointment to in_progress
    next_appt.queue_status = "in_progress"
    if next_appt.status == "requested":
        next_appt.status = "accepted"
    db.session.commit()

    socketio.emit("queue_update", {
        "event": "next",
        "appointment_id": next_appt.id,
        "queue_number": next_appt.queue_number
    })

    return jsonify({
        "msg": "Completed current patient and moved to next.",
        "completed": {
            "appointment_id": current.id,
            "queue_number": current.queue_number
        },
        "next": {
            "appointment_id": next_appt.id,
            "queue_number": next_appt.queue_number
        }
    })


@app.route("/queue/skip/<int:aid>", methods=["POST"])
@role_required(["doctor"])
def queue_skip(aid):
    appt = db.session.get(Appointment, aid)
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    appt.queue_status = "skipped"
    appt.status = "cancelled"
    db.session.commit()
    socketio.emit("queue_update", {
        "event": "skip",
        "appointment_id": appt.id
    })
    return jsonify({"msg": "Appointment skipped"})


@app.route("/queue/display", methods=["GET"])
def queue_display():
    """
    Public endpoint to show current token & upcoming patients (for TV screen).
    """
    today = datetime.utcnow().date()
    in_progress = db.session.query(Appointment).filter(
        func.date(Appointment.requested_time) == today,
        Appointment.queue_status == "in_progress"
    ).order_by(Appointment.queue_number.asc()).first()

    upcoming = db.session.query(Appointment).filter(
        func.date(Appointment.requested_time) == today,
        Appointment.queue_status == "waiting"
    ).order_by(Appointment.queue_number.asc()).limit(5).all()

    return jsonify({
        "current": {
            "id": in_progress.id,
            "queue_number": in_progress.queue_number
        } if in_progress else None,
        "upcoming": [
            {"id": a.id, "queue_number": a.queue_number}
            for a in upcoming
        ]
    })


# ----------- APPOINTMENT ANALYTICS ------------

@app.route("/analytics/appointments-per-week", methods=["GET"])
@role_required(["admin", "doctor"])
def appointments_per_week():
    data = db.session.query(
        func.yearweek(Appointment.requested_time, 3),
        func.count(Appointment.id)
    ).group_by(
        func.yearweek(Appointment.requested_time, 3)
    ).order_by(
        func.yearweek(Appointment.requested_time, 3)
    ).all()
    out = {
        "labels": [f"Week {row[0]}" for row in data],
        "dataset": [row[1] for row in data],
    }
    return jsonify(out)


@app.route("/analytics/patient-growth", methods=["GET"])
@role_required(["admin"])
def patient_growth():
    data = db.session.query(
        func.date(Appointment.requested_time),
        func.count(func.distinct(Appointment.patient_id))
    ).group_by(
        func.date(Appointment.requested_time)
    ).order_by(
        func.date(Appointment.requested_time)
    ).all()
    out = {
        "labels": [str(r[0]) for r in data],
        "dataset": [r[1] for r in data]
    }
    return jsonify(out)


@app.route("/analytics/revenue-per-month", methods=["GET"])
@role_required(["admin", "doctor"])
def revenue_per_month():
    data = db.session.query(
        func.date_format(Income.date, "%Y-%m"),
        func.sum(Income.amount)
    ).group_by(
        func.date_format(Income.date, "%Y-%m")
    ).order_by(
        func.date_format(Income.date, "%Y-%m")
    ).all()
    out = {
        "labels": [str(r[0]) for r in data],
        "dataset": [float(r[1]) if r[1] else 0.0 for r in data]
    }
    return jsonify(out)


# ----------- EXPORTS (EXCEL/ZIP) -----------

@app.route("/export/patients/xlsx", methods=["GET"])
@role_required(["admin"])
def export_patients():
    q = db.session.query(User).filter_by(role="patient")
    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Email", "Phone", "DOB", "Address"])
    for u in q.all():
        ws.append([
            u.id, u.name, u.email, u.phone,
            u.dob.isoformat() if u.dob else None,
            u.address
        ])
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="patients.xlsx")


@app.route("/export/appointments/xlsx", methods=["GET"])
@role_required(["admin"])
def export_appts():
    q = db.session.query(Appointment).all()
    wb = Workbook()
    ws = wb.active
    ws.append([
        "ID", "Patient", "Doctor", "Service",
        "Requested Time", "Status", "Payment Status",
        "Queue Number", "Visit Type"
    ])
    for a in q:
        ws.append([
            a.id,
            a.patient_id,
            a.doctor_id,
            a.service_id,
            a.requested_time.isoformat() if a.requested_time else None,
            a.status,
            a.payment_status,
            a.queue_number,
            a.visit_type
        ])
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="appointments.xlsx")


@app.route("/export/reports/zip", methods=["GET"])
@role_required(["admin"])
def export_reports_zip():
    q = db.session.query(Report).all()
    out = io.BytesIO()
    z = zipfile.ZipFile(out, "w")
    for r in q:
        if r.file_path and os.path.exists(r.file_path):
            z.write(r.file_path, arcname=r.filename)
    z.close()
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="reports.zip")


# ----------- PROFILE MANAGEMENT ------------

@app.route("/user/me", methods=["GET"])
@jwt_required()
def get_profile():
    user = db.session.get(User, int(get_jwt_identity()))
    if not user:
        return jsonify({"msg": "User not found"}), 404
    return jsonify({
        "id": user.id,
        "name": user.name,
        "email": user.email,
        "phone": user.phone,
        "dob": user.dob.isoformat() if user.dob else None,
        "address": user.address,
        "role": user.role,
        "profile_pic": user.profile_pic,
        "certificate_file": user.certificate_file
    })


@app.route("/user/update", methods=["POST"])
@jwt_required()
def update_profile():
    user = db.session.get(User, int(get_jwt_identity()))
    data = request.form if request.form else request.json or {}
    user.name = data.get("name", user.name)
    user.phone = data.get("phone", user.phone)
    user.address = data.get("address", user.address)
    if "dob" in data and data["dob"]:
        try:
            user.dob = datetime.strptime(data["dob"], "%Y-%m-%d").date()
        except Exception:
            pass

    if "profile_pic" in request.files:
        file = request.files["profile_pic"]
        if allowed_file(file.filename, "img"):
            fn = secure_filename(f"U{user.id}_" + file.filename)
            path = os.path.join(app.config["PROFILE_UPLOADS_DIR"], fn)
            file.save(path)
            user.profile_pic = path

    if user.role == "doctor" and "certificate_file" in request.files:
        file = request.files["certificate_file"]
        if allowed_file(file.filename, "cert"):
            fn = secure_filename(f"D{user.id}_cert_" + file.filename)
            path = os.path.join(app.config["CERT_UPLOADS_DIR"], fn)
            file.save(path)
            user.certificate_file = path

    db.session.commit()
    return jsonify({"msg": tr("profile_updated", parse_lang())})


# ----------- INVENTORY MANAGEMENT ------------

@app.route("/inventory", methods=["GET"])
@role_required(["admin", "doctor"])
def list_inventory():
    items = db.session.query(InventoryItem).all()
    return jsonify([
        {
            "id": i.id,
            "name": i.name,
            "quantity": i.quantity,
            "min_required": i.min_required,
            "expiry_date": i.expiry_date.isoformat() if i.expiry_date else None,
            "unit": i.unit
        } for i in items
    ])


@app.route("/inventory", methods=["POST"])
@role_required(["admin", "doctor"])
def create_or_update_inventory_item():
    data = request.json or {}
    item_id = data.get("id")
    name = data.get("name")
    if not name and not item_id:
        return jsonify({"msg": "name required for new item"}), 400

    if item_id:
        item = db.session.get(InventoryItem, _to_int(item_id))
        if not item:
            return jsonify({"msg": "Item not found"}), 404
    else:
        item = InventoryItem(name=name)
        db.session.add(item)

    if "name" in data:
        item.name = data["name"]
    if "quantity" in data:
        item.quantity = _to_int(data["quantity"]) or 0
    if "min_required" in data:
        item.min_required = _to_int(data["min_required"]) or 0
    if "unit" in data:
        item.unit = data["unit"]
    if "expiry_date" in data and data["expiry_date"]:
        try:
            item.expiry_date = datetime.strptime(data["expiry_date"], "%Y-%m-%d").date()
        except Exception:
            pass

    db.session.commit()
    return jsonify({"msg": tr("operation_success", parse_lang()), "id": item.id})


# ----------- RAZORPAY PAYMENTS ------------

@app.route("/payments/razorpay/create-order", methods=["POST"])
@role_required(["patient"])
def razorpay_create():
    if not (RAZORPAY_AVAILABLE and razorpay_client):
        return jsonify({"msg": "Razorpay not setup"}), 501
    data = request.json or {}
    appt_id = data.get("appointment_id")
    appt = db.session.get(Appointment, _to_int(appt_id))
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    service = db.session.get(Service, appt.service_id)
    if not service:
        return jsonify({"msg": "Service not found"}), 404
    amount = int(service.price) * 100
    order = razorpay_client.order.create({
        "amount": amount,
        "currency": "INR",
        "receipt": f"rcpt_{appt.id}",
        "payment_capture": 1
    })
    return jsonify({"order": order})


@app.route("/payments/razorpay/verify", methods=["POST"])
@role_required(["patient"])
def razorpay_verify():
    if not (RAZORPAY_AVAILABLE and razorpay_client):
        return jsonify({"msg": "Razorpay not setup"}), 501
    data = request.json or {}
    payment_id = data.get("razorpay_payment_id")
    order_id = data.get("razorpay_order_id")
    signature = data.get("razorpay_signature")
    appt_id = data.get("appointment_id")

    try:
        params_dict = {
            "razorpay_order_id": order_id,
            "razorpay_payment_id": payment_id,
            "razorpay_signature": signature
        }
        razorpay_client.utility.verify_payment_signature(params_dict)
    except Exception as ex:
        return jsonify({"msg": "Payment verification failed", "error": str(ex)}), 400

    appt = db.session.get(Appointment, _to_int(appt_id))
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404

    appt.payment_status = "paid"
    db.session.commit()
    ensure_income_for_appointment(appt)
    notify_event("Payment successful", appt)
    socketio.emit("payment_success", {"appointment_id": appt.id})
    return jsonify({"msg": "Payment successful", "id": appt.id})


# ----------- REPORTS (UPLOAD) -----------

@app.route("/reports/upload", methods=["POST"])
@role_required(["doctor"])
def upload_report():
    if "file" not in request.files or "appointment_id" not in request.form:
        return jsonify({"msg": "file and appointment_id required"}), 400
    file = request.files["file"]
    appt_id_raw = request.form["appointment_id"]
    appt_id = _to_int(appt_id_raw)
    if appt_id is None:
        return jsonify({"msg": "invalid appointment_id"}), 400
    appt = db.session.get(Appointment, appt_id)
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    if appt.doctor_id != int(get_jwt_identity()):
        return jsonify({"msg": "Forbidden: not your appointment"}), 403
    raw_filename = file.filename or "report"
    safe_name = secure_filename(raw_filename)

    file_path = os.path.join(app.config["REPORTS_DIR"], safe_name)
    file.save(file_path)

    r = Report(appointment_id=appt_id, filename=safe_name, file_path=file_path)
    db.session.add(r)
    db.session.commit()
    try:
        socketio.emit("report_uploaded", {
            "appointment_id": appt_id,
            "report_id": r.id,
            "uploaded_at": r.uploaded_at.isoformat(),
        })
    except Exception:
        pass
    notify_event("Report uploaded", appt)
    return jsonify({"msg": "Report uploaded!", "id": r.id, "filename": safe_name}), 201


# ----------- PRESCRIPTIONS (DIGITAL) -----------

@app.route("/prescriptions", methods=["POST"])
@role_required(["doctor"])
def create_prescription():
    data = request.json or {}
    appointment_id = data.get("appointment_id")
    diagnosis = data.get("diagnosis")
    notes = data.get("notes")
    medicines = data.get("medicines") or []

    appt = db.session.get(Appointment, _to_int(appointment_id))
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404

    doctor_id = int(get_jwt_identity())
    if appt.doctor_id != doctor_id:
        return jsonify({"msg": "Not your appointment"}), 403

    presc = Prescription(
        appointment_id=appt.id,
        doctor_id=doctor_id,
        patient_id=appt.patient_id,
        diagnosis=diagnosis,
        notes=notes,
        medicines=medicines
    )
    db.session.add(presc)
    db.session.commit()

    generate_prescription_pdf(presc)

    socketio.emit("prescription_created", {
        "appointment_id": appt.id,
        "prescription_id": presc.id,
        "created_at": presc.created_at.isoformat(),
    })

    return jsonify({
        "msg": "Prescription created",
        "id": presc.id
    }), 201


@app.route("/prescriptions/<int:pid>", methods=["GET"])
@jwt_required()
def get_prescription(pid):
    presc = db.session.get(Prescription, pid)
    if not presc:
        return jsonify({"msg": "Prescription not found"}), 404

    user_id = int(get_jwt_identity())
    claims = get_jwt()
    role = claims.get("role")

    if role == "patient" and presc.patient_id != user_id:
        return jsonify({"msg": "Forbidden"}), 403
    if role == "doctor" and presc.doctor_id != user_id:
        return jsonify({"msg": "Forbidden"}), 403

    return jsonify({
        "id": presc.id,
        "appointment_id": presc.appointment_id,
        "doctor_id": presc.doctor_id,
        "patient_id": presc.patient_id,
        "diagnosis": presc.diagnosis,
        "notes": presc.notes,
        "medicines": presc.medicines,
        "created_at": presc.created_at.isoformat(),
        "has_pdf": bool(presc.pdf_path)
    })


@app.route("/prescriptions/appointment/<int:aid>", methods=["GET"])
@jwt_required()
def list_prescriptions_for_appointment(aid):
    appt = db.session.get(Appointment, aid)
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404

    user_id = int(get_jwt_identity())
    claims = get_jwt()
    role = claims.get("role")

    if role == "patient" and appt.patient_id != user_id:
        return jsonify({"msg": "Forbidden"}), 403
    if role == "doctor" and appt.doctor_id != user_id:
        return jsonify({"msg": "Forbidden"}), 403

    prescs = db.session.query(Prescription).filter_by(appointment_id=aid).all()
    return jsonify([
        {
            "id": p.id,
            "diagnosis": p.diagnosis,
            "created_at": p.created_at.isoformat(),
            "has_pdf": bool(p.pdf_path)
        } for p in prescs
    ])


@app.route("/prescriptions/<int:pid>/pdf", methods=["GET"])
@jwt_required()
def download_prescription_pdf(pid):
    presc = db.session.get(Prescription, pid)
    if not presc:
        return jsonify({"msg": "Prescription not found"}), 404

    user_id = int(get_jwt_identity())
    claims = get_jwt()
    role = claims.get("role")

    if role == "patient" and presc.patient_id != user_id:
        return jsonify({"msg": "Forbidden"}), 403
    if role == "doctor" and presc.doctor_id != user_id:
        return jsonify({"msg": "Forbidden"}), 403

    if not presc.pdf_path or not os.path.exists(presc.pdf_path):
        return jsonify({"msg": "PDF not generated"}), 404

    return send_file(presc.pdf_path, as_attachment=True, download_name=os.path.basename(presc.pdf_path))


# ----------- PATIENT VISIT HISTORY TIMELINE -----------

@app.route("/patients/<int:pid>/timeline", methods=["GET"])
@jwt_required()
def patient_timeline(pid):
    user_id = int(get_jwt_identity())
    claims = get_jwt()
    role = claims.get("role")

    if role == "patient" and user_id != pid:
        return jsonify({"msg": "Forbidden"}), 403

    if role not in ("patient", "doctor", "admin"):
        return jsonify({"msg": "Forbidden"}), 403

    appts = db.session.query(Appointment).filter_by(patient_id=pid).order_by(Appointment.requested_time.asc()).all()
    items = []

    for a in appts:
        ts = a.requested_time.isoformat() if a.requested_time else None
        service = db.session.get(Service, a.service_id) if a.service_id else None

        items.append({
            "type": "appointment",
            "timestamp": ts,
            "appointment_id": a.id,
            "status": a.status,
            "queue_number": a.queue_number,
            "visit_type": a.visit_type,
            "service": service.name if service else None
        })

        for r in a.reports:
            items.append({
                "type": "report",
                "timestamp": r.uploaded_at.isoformat() if r.uploaded_at else ts,
                "appointment_id": a.id,
                "report_id": r.id,
                "filename": r.filename
            })

        for p in a.prescriptions:
            items.append({
                "type": "prescription",
                "timestamp": p.created_at.isoformat() if p.created_at else ts,
                "appointment_id": a.id,
                "prescription_id": p.id,
                "diagnosis": p.diagnosis
            })

        for inc in a.income_entries:
            items.append({
                "type": "payment",
                "timestamp": inc.date.isoformat(),
                "appointment_id": a.id,
                "amount": inc.amount
            })

    items_sorted = sorted(items, key=lambda x: x["timestamp"] or "")
    return jsonify(items_sorted)


# ----------- FINANCE: INCOME & EXPENSE TRACKER -----------

@app.route("/finance/expenses", methods=["POST"])
@role_required(["admin", "doctor"])
def add_expense():
    data = request.json or {}
    amount = data.get("amount")
    if amount is None:
        return jsonify({"msg": "amount required"}), 400
    try:
        amount = float(amount)
    except Exception:
        return jsonify({"msg": "invalid amount"}), 400

    exp_date = data.get("date")
    if exp_date:
        try:
            d = datetime.strptime(exp_date, "%Y-%m-%d").date()
        except Exception:
            d = date.today()
    else:
        d = date.today()

    exp = Expense(
        date=d,
        type=data.get("type") or "Other",
        amount=amount,
        notes=data.get("notes")
    )
    db.session.add(exp)
    db.session.commit()
    return jsonify({"msg": "Expense added", "id": exp.id}), 201


@app.route("/finance/summary", methods=["GET"])
@role_required(["admin", "doctor"])
def finance_summary():
    from_str = request.args.get("from")
    to_str = request.args.get("to")
    today = date.today()
    if to_str:
        try:
            to_d = datetime.strptime(to_str, "%Y-%m-%d").date()
        except Exception:
            to_d = today
    else:
        to_d = today

    if from_str:
        try:
            from_d = datetime.strptime(from_str, "%Y-%m-%d").date()
        except Exception:
            from_d = to_d - timedelta(days=30)
    else:
        from_d = to_d - timedelta(days=30)

    total_income = db.session.query(func.sum(Income.amount)).filter(Income.date.between(from_d, to_d)).scalar() or 0.0
    total_expense = db.session.query(func.sum(Expense.amount)).filter(Expense.date.between(from_d, to_d)).scalar() or 0.0

    return jsonify({
        "from": from_d.isoformat(),
        "to": to_d.isoformat(),
        "total_income": total_income,
        "total_expense": total_expense,
        "profit": total_income - total_expense
    })


# ----------- DOCTOR DASHBOARD (CLINIC VIEW) -----------

@app.route("/dashboard/doctor", methods=["GET"])
@role_required(["doctor"])
def doctor_dashboard():
    doctor_id = int(get_jwt_identity())
    today = datetime.utcnow().date()
    start_month = today.replace(day=1)
    one_week_ago = today - timedelta(days=7)

    today_appts = db.session.query(Appointment).filter(
        Appointment.doctor_id == doctor_id,
        func.date(Appointment.requested_time) == today,
        Appointment.status != "cancelled"
    ).count()

    pending_payments = db.session.query(Appointment).filter(
        Appointment.doctor_id == doctor_id,
        Appointment.payment_status != "paid",
        Appointment.status.in_(["accepted", "completed"])
    ).count()

    monthly_revenue = db.session.query(func.sum(Income.amount)).filter(
        Income.date.between(start_month, today)
    ).scalar() or 0.0

    new_patients_this_week = db.session.query(
        func.count(func.distinct(Appointment.patient_id))
    ).filter(
        Appointment.doctor_id == doctor_id,
        func.date(Appointment.requested_time).between(one_week_ago, today)
    ).scalar() or 0

    walkins_today = db.session.query(Appointment).filter(
        Appointment.doctor_id == doctor_id,
        Appointment.visit_type == "walkin",
        func.date(Appointment.requested_time) == today
    ).count()

    thirty_days_ago = today - timedelta(days=30)
    svc_rows = db.session.query(
        Service.name,
        func.count(Appointment.id).label("cnt")
    ).join(Appointment, Appointment.service_id == Service.id).filter(
        Appointment.doctor_id == doctor_id,
        func.date(Appointment.requested_time).between(thirty_days_ago, today)
    ).group_by(
        Service.id
    ).order_by(
        func.count(Appointment.id).desc()
    ).limit(5).all()
    top_services = [{"name": r[0], "count": r[1]} for r in svc_rows]

    low_stock = db.session.query(InventoryItem).filter(
        InventoryItem.quantity <= InventoryItem.min_required
    ).all()
    expiring_items = db.session.query(InventoryItem).filter(
        InventoryItem.expiry_date.isnot(None),
        InventoryItem.expiry_date <= (today + timedelta(days=30))
    ).all()

    return jsonify({
        "today_appointments": today_appts,
        "pending_payments": pending_payments,
        "monthly_revenue": monthly_revenue,
        "new_patients_this_week": new_patients_this_week,
        "walkins_today": walkins_today,
        "top_services": top_services,
        "low_stock": [
            {
                "id": i.id,
                "name": i.name,
                "quantity": i.quantity,
                "min_required": i.min_required,
                "unit": i.unit
            } for i in low_stock
        ],
        "expiring_items": [
            {
                "id": i.id,
                "name": i.name,
                "expiry_date": i.expiry_date.isoformat() if i.expiry_date else None,
                "quantity": i.quantity,
                "unit": i.unit
            } for i in expiring_items
        ]
    })


# ----------- SERVICES LIST ------------

@app.route("/services", methods=["GET"])
def list_services():
    services = db.session.query(Service).all()
    return jsonify([{"id": s.id, "name": s.name, "price": s.price} for s in services])


# ----------- STRIPE WEBHOOK ------------

@app.route("/stripe/webhook", methods=["POST"])
def stripe_webhook():
    if not STRIPE_AVAILABLE:
        return jsonify({"msg": "Stripe not available on server"}), 501

    payload = request.data
    sig_header = request.headers.get("stripe-signature")
    try:
        if STRIPE_WEBHOOK_SECRET and sig_header:
            event = stripe.Webhook.construct_event(
                payload=payload,
                sig_header=sig_header,
                secret=STRIPE_WEBHOOK_SECRET
            )
        else:
            event = stripe.Event.construct_from(request.get_json(), stripe.api_key)
    except Exception as e:
        return jsonify({"msg": "Webhook error", "error": str(e)}), 400

    etype = getattr(event, "type", None) or event.get("type")
    if etype in ("payment_intent.succeeded", "checkout.session.completed"):
        data = event.get("data", {}).get("object", {})
        appointment_id = None
        if isinstance(data, dict):
            appointment_id = data.get("metadata", {}).get("appointment_id") or data.get("client_reference_id")
        if appointment_id:
            appt = db.session.get(Appointment, _to_int(appointment_id))
            if appt:
                appt.payment_status = "paid"
                db.session.commit()
                ensure_income_for_appointment(appt)
                socketio.emit("payment_success", {"appointment_id": appt.id})
    return jsonify({"msg": "received"}), 200


# ----------- APPOINTMENT STATUS OPERATIONS ------------

@app.route("/appointments/<id>/accept", methods=["POST"])
@role_required(["doctor"])
def accept_appointment(id):
    appt = db.session.get(Appointment, _to_int(id))
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    appt.status = "accepted"
    appt.meeting_link = f"https://meet.jit.si/{uuid.uuid4()}"
    appt.queue_status = appt.queue_status or "waiting"
    db.session.commit()
    notify_event("Appointment accepted", appt)
    socketio.emit("appointment_update", {
        "id": appt.id,
        "status": "accepted",
        "meeting_link": appt.meeting_link
    })
    return jsonify({"msg": "Accepted", "meeting_link": appt.meeting_link})


@app.route("/appointments/<id>/complete", methods=["POST"])
@role_required(["doctor"])
def complete_appointment(id):
    appt = db.session.get(Appointment, _to_int(id))
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    appt.status = "completed"
    appt.queue_status = "done"
    db.session.commit()
    socketio.emit("appointment_update", {"id": appt.id, "status": "completed"})
    return jsonify({"msg": "Appointment completed"})


@app.route("/appointments/<id>/reschedule", methods=["POST"])
@role_required(["patient"])
def reschedule_appointment(id):
    """
    Option A reschedule:
      - Same rules as initial booking:
          * 09:00–18:00 IST
          * no double booking for same doctor/time
    """
    data = request.json or {}
    new_time_str = data.get("new_time")
    if not new_time_str:
        return jsonify({"msg": "new_time required"}), 400

    appt = db.session.get(Appointment, _to_int(id))
    if not appt or appt.patient_id != int(get_jwt_identity()):
        return jsonify({"msg": "Appointment not found or unauthorized"}), 404

    try:
        dt_ist, dt_utc = parse_client_datetime(new_time_str)
    except Exception:
        return jsonify({"msg": "Invalid time format"}), 400

    # 9–18 check in IST
    if dt_ist.hour < 9 or dt_ist.hour >= 18:
        return jsonify({"msg": "Reschedule allowed only between 09:00 and 18:00"}), 400

    # Prevent double booking for same doctor/time
    conflict = db.session.query(Appointment).filter(
        Appointment.doctor_id == appt.doctor_id,
        Appointment.requested_time == dt_utc,
        Appointment.id != appt.id,
        Appointment.status != "cancelled"
    ).first()
    if conflict:
        return jsonify({"msg": "New time already booked"}), 400

    appt.requested_time = dt_utc
    appt.status = "rescheduled"
    appt.queue_status = "waiting"
    db.session.commit()
    notify_event("Appointment rescheduled", appt)
    socketio.emit("appointment_update", {
        "id": appt.id,
        "status": "rescheduled",
        "requested_time": dt_utc.isoformat()
    })
    return jsonify({"msg": "Appointment rescheduled"})


@app.route("/appointments/<id>/cancel", methods=["POST"])
@role_required(["patient", "doctor"])
def cancel_appointment(id):
    appt = db.session.get(Appointment, _to_int(id))
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    user_id = int(get_jwt_identity())
    user = db.session.get(User, user_id)
    if user.role not in ("admin",) and user_id not in (appt.patient_id, appt.doctor_id):
        return jsonify({"msg": "Unauthorized"}), 403
    appt.status = "cancelled"
    appt.queue_status = "skipped"
    db.session.commit()
    notify_event("Appointment cancelled", appt)
    socketio.emit("appointment_update", {"id": appt.id, "status": "cancelled"})
    return jsonify({"msg": "Appointment cancelled"})


@app.route("/appointments/<int:aid>/bill", methods=["GET"])
@jwt_required()
def bill(aid):
    appt = db.session.get(Appointment, aid)
    if not appt:
        return jsonify({"msg": "Appointment not found"}), 404
    service = db.session.get(Service, appt.service_id)
    user_id = int(get_jwt_identity())
    claims = get_jwt()
    role = claims.get("role")

    if not ((role == "doctor" and appt.doctor_id == user_id) or (role == "patient" and appt.patient_id == user_id)):
        return jsonify({"msg": "Forbidden"}), 403

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    y = 720
    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, y, safe_text_for_pdf("OralCare System - Bill"))
    y -= 30
    c.setFont("Helvetica", 12)
    c.drawString(100, y, safe_text_for_pdf(f"Appointment ID: {appt.id}"))
    y -= 20
    if service:
        c.drawString(100, y, safe_text_for_pdf(f"Service: {service.name}"))
        y -= 20
        c.drawString(100, y, safe_text_for_pdf(f"Price: ₹{service.price}"))
        y -= 20
    else:
        c.drawString(100, y, safe_text_for_pdf("Service: n/a"))
        y -= 20
    c.drawString(100, y, safe_text_for_pdf(f"Payment Status: {appt.payment_status or 'Pending'}"))
    c.showPage()
    c.save()
    buf.seek(0)
    return send_file(buf, download_name="bill.pdf", as_attachment=True)


# ----------- HOME ------------

@app.route("/")
def home():
    try:
        return render_template("index.html")
    except Exception:
        return "<h3>OralCare Backend</h3><p>Clinic Edition API is running.</p>"


# ----------- MAIN ------------

if __name__ == "__main__":
    print("🚀 OralCare backend  running on http://127.0.0.1:5000")
    socketio.run(app, host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=True)