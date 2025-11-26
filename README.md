🦷 OralCare  System

A complete single-doctor clinic management system built with:

Flask (Python) — Backend API

MySQL — Database

HTML / CSS / JS — Frontend

JWT Authentication

Socket.IO — Real-time Queue Updates

Email & SMS Notifications

Razorpay Integration (optional)

This project manages appointments, walk-ins, queue system, prescriptions, reports, inventory, finance, admin panel, and more.

🚀 Features
👨‍⚕️ Single-Doctor Clinic

✔ Only one doctor in the clinic
✔ Simplified appointment booking (9 AM – 6 PM)
✔ Auto increments queue number

🧑‍💻 Authentication

✔ Login with email + password
✔ Login with OTP (Email)
✔ Admin Login
✔ JWT-based sessions
✔ Email notifications

🗓️ Appointments
Patient

✔ Book appointment with service, date & time
✔ View appointment history
✔ View prescriptions, bills & reports
✔ Track timeline (all visits)

Doctor

✔ Today's queue
✔ Call next patient
✔ Skip patient
✔ Complete current & call next
✔ Create prescription
✔ Upload/view reports

Queue System

✔ Real-time updates using Socket.IO
✔ Public TV queue display
✔ Automatic status changes:

waiting → in_progress → done/skip
✔ Email/SMS for:

booked

completed (optional)

🧾 Prescriptions

✔ Add diagnosis, notes
✔ Add unlimited medicines
✔ Auto-save with DB
✔ PDF download
✔ Shown in timeline

📄 Reports

✔ Upload any file (X-ray, lab report, etc.)
✔ Linked to appointment
✔ Visible in patient timeline
✔ Download as needed

💰 Payments

✔ Razorpay order creation endpoint
✔ Display pending payments

📦 Inventory

✔ Add / update stock
✔ Track low stock (auto-detect)
✔ Expiry dates
✔ Doctor/Admin both can manage

💸 Finance

✔ Add expenses
✔ Summary for last 30 days or date range
✔ Calculates:

total income

total expense

profit

🛠 Admin Panel

✔ Create doctors
✔ Manage users
✔ Analytics:

Revenue per month

Appointments per week
✔ Export:

Patients (XLSX)

Appointments (XLSX)

Reports (ZIP)

🏗️ Tech Stack
Layer	Technology
Backend	Flask, SQLAlchemy, JWT, MySQL
Frontend	HTML, CSS,  JS
Realtime	Socket.IO
Database	MySQL 8+
Charts	Chart.js
Auth	JWT + OTP
Notifications	SMTP (Email), SMS (optional)
📡 API Overview (Short Summary)
Authentication

POST /login
POST /auth/send-otp
POST /auth/verify-otp

Appointments

POST /appointments/book
POST /appointments/walkin
GET /queue/today
POST /queue/next
POST /queue/skip/<id>
POST /queue/complete-and-next
GET /queue/display

Prescriptions

POST /prescriptions
GET /prescriptions/appointment/<id>
GET /prescriptions/<id>/pdf

Reports

POST /reports/upload

Payments

POST /payments/razorpay/create-order

Inventory

POST /inventory
GET /inventory

Finance

POST /finance/expenses
GET /finance/summary

Admin

GET /admin/users
POST /admin/create-doctor
DELETE /admin/delete-user/<id>

🛑 Setup Instructions
1. Clone repository
git clone https://github.com/<your-repo>/oralcare.git
cd oralcare

2. Create Virtual Environment
python -m venv venv
venv/Scripts/activate   # Windows
source venv/bin/activate # Linux/Mac

3. Install Dependencies
pip install -r requirements.txt

4. Configure Environment (.env)
DATABASE_URI = mysql+pymysql://root:<password>@localhost:3306/oralcare_db
JWT_SECRET_KEY = your_jwt_secret
EMAIL_HOST = smtp.gmail.com
EMAIL_PORT = 587
EMAIL_USER = youremail@gmail.com
EMAIL_PASS = your_app_password

5. Run Migrations
flask shell
>>> from app import db
>>> db.create_all()
>>> exit()

6. Run Backend
python app.py

Default URL → http://localhost:5000
