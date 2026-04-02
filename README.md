# Lab Attendance System

A Flask-based attendance tracking system for university lab sessions.
Supports USB barcode scanners, phone camera scanning, time-in/time-out logging,
duplicate scan protection, and semester attendance reports with 75% threshold flagging.

---

## Features

- **Login** — hashed password auth, admin and staff roles
- **Multiple courses** — each with its own student roster
- **CSV roster upload** — import students from a spreadsheet export
- **USB barcode scanning** — plug in any USB HID barcode scanner; it acts as a keyboard
- **Camera scanning** — use a phone or laptop camera via ZXing (requires HTTPS)
- **Time-in / time-out** — first scan = in, second scan = out
- **Duplicate protection** — same ID scanned twice within 30s is ignored
- **Session CSV export** — per-session attendance download
- **Semester Excel report** — colour-coded, flags students below 75%

---

## Running locally

```bash
# 1. Clone / download the project
cd labattend

# 2. Create virtual environment
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run
python app.py
```

Open http://localhost:5000

Default login: **admin** / **changeme123**
(Change this immediately via the Users page.)

> **Note:** Camera scanning requires HTTPS. On localhost it only works
> in Chrome/Edge (which exempt localhost from the HTTPS requirement).
> On mobile, deploy to Render first.

---

## Deploying to Render (free, with persistent PostgreSQL)

1. Push your code to a GitHub repository (public or private).

2. Go to https://render.com → New → Blueprint → connect your repo.
   Render reads `render.yaml` and automatically creates:
   - A free PostgreSQL database (`lab-attendance-db`)
   - A web service connected to it via `DATABASE_URL`

3. In the Render dashboard, set one environment variable manually:
   - `ADMIN_PASSWORD` — your chosen admin password

   `SECRET_KEY` is auto-generated. `DATABASE_URL` is wired automatically
   from the database. `DEDUP_SECONDS` defaults to 30.

4. Deploy. Your app gets a free HTTPS URL, e.g. `lab-attendance.onrender.com`.
   The PostgreSQL database persists indefinitely — data survives redeploys.

> **Render free tier caveat:** The web service "sleeps" after 15 minutes of
> inactivity and takes ~30 seconds to wake on the next request. For a live lab
> session this is fine once it's awake. The free PostgreSQL instance expires
> after 90 days — Render will email you; just create a new one and update
> `DATABASE_URL`. Upgrade to the $7/month plan for always-on.

> **Local development with PostgreSQL:** Install PostgreSQL locally and set
> `DATABASE_URL=postgresql://localhost/labattend` before running `python app.py`.
> Or keep a `.env` file and use `python-dotenv` to load it.

---

## CSV roster format

Your CSV must have at minimum two columns. Column names are flexible —
the app looks for any column containing "id" and any column containing "name".

```
student_id,name
816001001,Alicia Ramkhelawan
816001002,Brian Seepersad
816001003,Candice Mohammed
```

You can export this directly from Banner, SIS, or a spreadsheet.
Save as CSV (UTF-8) before uploading.

---

## USB barcode scanner setup

1. Plug the scanner into the laptop via USB.
2. Open the scan session page in a browser.
3. Click the "USB / Keyboard scanner" button (active by default).
4. The input field is auto-focused — scan a student's ID card.
5. The scanner sends the barcode value + Enter automatically.

No drivers or configuration needed. USB HID barcode scanners emulate a keyboard.

---

## Camera scanner (phone)

1. Deploy to Render (HTTPS required for camera access on phones).
2. Open the session page on your phone browser.
3. Click "Camera scanner" — grant camera permission when prompted.
4. Hold a student's ID card steady in front of the camera.
5. ZXing decodes the barcode automatically within 1–2 seconds.

Works best in good lighting. Rear camera is preferred automatically.

Supported barcode formats: Code 128, Code 39, QR Code, EAN, UPC, PDF417, and more.

---

## Adding staff users

1. Log in as admin.
2. Go to **Users** in the navigation bar.
3. Create a new user with a username (their UWI ID or a short name) and password.
4. Share the credentials securely (e.g. in person or via encrypted message).

Staff can scan attendance and view reports but cannot manage users.

---

## Adjusting the duplicate scan window

Set the `DEDUP_SECONDS` environment variable (default: 30).
A scan of the same student ID within this window is silently ignored,
preventing accidental double-scans when the barcode is read twice in rapid succession.

---

## Adding fingerprint scanner support (future)

When a USB fingerprint scanner is added:

1. Install the vendor SDK or `pyfingerprint` library on the server machine.
2. Add a `fingerprint_id` column to the `students` table.
3. Create an enrollment route that maps a fingerprint template to a student ID.
4. Create a background thread that polls the scanner and calls the same
   `/api/scan` endpoint with the resolved student ID.

The rest of the system (dedup, time-in/out, reports) remains unchanged.

---

## File structure

```
labattend/
├── app.py                  # Flask app, routes, API, exports
├── requirements.txt
├── render.yaml             # Render deployment config
├── attendance.db           # SQLite database (auto-created on first run)
└── templates/
    ├── base.html           # Shared layout, nav, styles
    ├── login.html
    ├── courses.html
    ├── new_course.html
    ├── course_detail.html
    ├── upload_roster.html
    ├── scan_session.html   # Main scanning UI (USB + camera)
    ├── report.html         # Semester attendance report
    └── admin_users.html
```
