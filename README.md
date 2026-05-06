# Job Application Tracker — Setup Guide

# 💼 Job Application Tracker

A Windows desktop app to track job applications — built with Python and Tkinter.
Connects directly to an Excel file. No database, no internet, works 100% offline.

## Features
- ➕ Add applications with auto-assigned ID and date picker
- 🔍 Search applications by date
- 📊 View all applications in a live table — click any row to Edit or Remove
- 📈 Analytics dashboard with counts by website and recruitment stage
- 💾 All changes save directly to Excel — no import/export ever

## Screenshots
<!-- Add screenshots here after uploading them -->

## Tech Stack
- Python 3.x
- Tkinter (GUI)
- openpyxl (Excel read/write)
- PyInstaller (builds to .exe)

## How to Run

### Option A — Run with Python
```bash
pip install openpyxl
python app.py
```

### Option B — Build as .exe

## What's in this folder

| File | Purpose |
|------|---------|
| `launcher.py` | Entry point — auto-connects to Excel on first launch |
| `app.py` | Main application code |
| `build.bat` | One-click script to build `JobApplicationTracker.exe` |
| `Job-application-tracksheet.xlsx` | Excel data file |

---

## Step 1 — Install Python (one time only)

1. Go to **https://python.org/downloads**
2. Download Python **3.9 or newer**
3. ✅ During install, **check "Add Python to PATH"** — this is important!
4. Click Install Now

---

## Step 2 — Build the EXE

1. Put all 4 files above **in the same folder** on your PC
2. Double-click **`build.bat`**
3. If build.bat crashes or does not work then run diagons.bat
3. A black window opens — wait 1–2 minutes
4. When done you'll see a **`dist\`** folder appear

Inside `dist\` you'll find:
- `JobApplicationTracker.exe`
- `Job-application-tracksheet.xlsx`

---

## Step 3 — Run the app

1. Open the `dist\` folder
2. Double-click **`JobApplicationTracker.exe`**
3. ✅ The app opens **already connected** to your Excel file — no setup needed

---

## Rules to keep it working

> **Always keep `JobApplicationTracker.exe` and `Job-application-tracksheet.xlsx` in the same folder.**

That's the only rule. Every edit you make in the app saves directly to the Excel file. No importing, no exporting.

If you ever move the app, move **both files** together.

---

## How each section works

### ➕ Add Application
- Application ID is assigned automatically
- Date defaults to today — change with the calendar or type DD-MM-YYYY
- Call Received, Interview, Rounds, Rejected all default to **No** — only change if needed
- Click **Submit** to save — data goes straight to Excel

### 🔍 Search by Date
- Enter any date (DD-MM-YYYY) or pick from the calendar
- Results show all applications for that date
- Click **Edit** to modify a record — confirm to save
- Click **Remove** to delete — you'll be warned before it's gone

### 📊 All Applications
- Shows your full Excel sheet live
- Automatically refreshes when you switch to this tab

### 📈 Analytics
- KPI tiles: total apps, calls, interviews, rejections
- Bar charts: breakdown by website and by recruitment stage

---

## Troubleshooting

**App asks to connect every time (on Python/script version):**
Make sure `config.json` is in the same folder as `app.py` — it stores your file path.

**EXE asks to connect every time:**
Make sure `Job-application-tracksheet.xlsx` is in the **same folder** as `JobApplicationTracker.exe`.

**Excel file not updating:**
Make sure the Excel file is **not open in Microsoft Excel** at the same time as the tracker app. Close Excel first.

**"Python not found" when running build.bat:**
Re-install Python from python.org and make sure you ticked **"Add Python to PATH"** during install.
