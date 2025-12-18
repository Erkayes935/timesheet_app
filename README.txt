Timesheet App - Final Version
==============================

Description:
A lightweight Python Tkinter application for managing timesheet entries with SQLite persistence,
Excel export, and Google Sheets synchronization capabilities.

Tech Stack:
- Python 3.8+
- Tkinter (GUI)
- SQLite (Database)
- openpyxl (Excel export)
- gspread (Google Sheets API)
- oauth2client (Google authentication)

How to run (Windows):
1. Ensure Python 3.8+ is installed.
2. Install dependencies: pip install -r requirements.txt
3. (Optional) If using Google Sheets sync, set up credentials.json with your Google API credentials.
4. Copy all files (main.py, google_sheet_sync.py, credentials.json) to a folder.
5. Run: python main.py
6. The app will create `timesheet.db` in the same folder.

Features:
- Track work hours with dual time slots (jam_mulai_1/jam_selesai_1, jam_mulai_2/jam_selesai_2)
- Track overtime hours (lembur_mulai, lembur_selesai) with reasons and descriptions
- 24-hour format time pickers in 5-minute increments
- Autocomplete suggestions for overtime reasons
- Calendar view to navigate entries by date
- Export all entries to Excel (.xlsx file)
- Sync daily entries to Google Sheets (creates separate sheet per date)
- SQLite database stores all entries with automatic schema patches
- Indonesian language support (Identitas, Tanggal, Deskripsi, etc.)

Building to .exe (optional):
- Install pyinstaller: pip install pyinstaller
- Then run:
    pyinstaller --onefile --noconsole main.py
- The built exe will be in the `dist` folder.
- The app is self-contained and does not require localhost or internet (unless using Google Sheets sync).

Database Schema:
The SQLite database includes the following columns:
- id, entry_date, jam_mulai_1, jam_selesai_1, jam_mulai_2, jam_selesai_2
- lembur_mulai, lembur_selesai, alasan_lembur, deskripsi_lembur, note

Google Sheets Integration:
- Requires credentials.json file with Google Service Account credentials
- Each day's entry is synced to a separate worksheet
- Includes calculated totals for work hours and overtime