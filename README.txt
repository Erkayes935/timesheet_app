Timesheet Lightweight (Python + Tkinter + SQLite + openpyxl)
----------------------------------------------------------

How to run (Windows):
1. Ensure Python 3.8+ is installed.
2. Install dependency: pip install openpyxl
3. Copy main.py to a folder where you want the app.
4. Run: python main.py
5. The app will create `timesheet.db` in the same folder.

Building to .exe (optional):
- Install pyinstaller: pip install pyinstaller
- Then run:
    pyinstaller --onefile --noconsole main.py
- The built exe will be in the `dist` folder. The app is self-contained and does not require a localhost.

Notes:
- The UI uses 24-hour format pickers in 5-minute increments.
- You can export all saved entries to an .xlsx file.
- The SQLite database `timesheet.db` stores all entries per date.

If you want, I can also build the exe here and provide the downloadable binary (but it may be large).