import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, date, timedelta
import calendar
from openpyxl import Workbook
from google_sheet_sync import GoogleSheetSync
import os
import sys

# =============================================================
# RESOURCE PATH (FOR PYINSTALLER)
# =============================================================
def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller exe
    """
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

CREDENTIALS_PATH = resource_path("credentials.json")

# =============================================================
# DATABASE PATH
# =============================================================
def app_data_path(filename):
    base_dir = os.path.join(
        os.environ.get("APPDATA") or os.path.expanduser("~"),
        "TimesheetApp"
    )
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, filename)

DB_PATH = "timesheet.db"


# =============================================================
# DATABASE PATCH
# =============================================================
def ensure_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS entries (
            id INTEGER PRIMARY KEY,
            entry_date TEXT NOT NULL,
            jam_mulai_1 TEXT,
            jam_selesai_1 TEXT,
            jam_mulai_2 TEXT,
            jam_selesai_2 TEXT,
            lembur_mulai TEXT,
            lembur_selesai TEXT,
            alasan_lembur TEXT,
            deskripsi_lembur TEXT,
            note TEXT
        )
    """)

    # patch kolom jika belum ada
    try: cur.execute("ALTER TABLE entries ADD COLUMN alasan_lembur TEXT")
    except: pass
    try: cur.execute("ALTER TABLE entries ADD COLUMN deskripsi_lembur TEXT")
    except: pass

    conn.commit()
    conn.close()


# =============================================================
# UTILITAS WAKTU
# =============================================================
def to_minutes(t: str):
    if not t:
        return 0
    try:
        h, m = map(int, t.split(":"))
        return h * 60 + m
    except:
        return 0


def format_duration(minutes: int):
    if minutes <= 0:
        return "0 menit"
    h = minutes // 60
    m = minutes % 60
    return f"{h} jam {m} menit"


def calc_total_kerja(jm1, js1, jm2, js2):
    total = 0

    a = to_minutes(jm1)
    b = to_minutes(js1)
    c = to_minutes(jm2)
    d = to_minutes(js2)

    if b > a:
        total += (b - a)
    if d > c:
        total += (d - c)

    return format_duration(total)


def calc_total_lembur(lm, ls):
    a = to_minutes(lm)
    b = to_minutes(ls)

    if a == 0 or b == 0:
        return "0 menit"

    if b < a:
        b += 24 * 60   # lembur lewat tengah malam

    dur = b - a
    return format_duration(dur)


ALASAN_PRESET = [
    "Deadline proyek",
    "Bug fix urgent",
    "Support implementasi",
    "Maintenance sistem",
    "Lainnya...",
]


# =============================================================
# APLIKASI UTAMA
# =============================================================
class TimesheetApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Timesheet App Final Version")
        self.root.geometry("1200x750")

        ensure_db()

        self.gs = None  # Google sheet handler

        self.build_ui()
        self.load_month()

    # =========================================================
    # TIME LIST
    # =========================================================
    def make_time_values(self, step=5):
        vals = []
        for h in range(24):
            for m in range(0, 60, step):
                vals.append(f"{h:02d}:{m:02d}")
        return vals

    # =========================================================
    # AUTOCOMPLETE
    # =========================================================
    def enable_autocomplete(self, cb, values):
        def on_key(event):
            typed = cb.get()
            matches = [v for v in values if v.startswith(typed)]
            cb["values"] = matches if matches else values
        cb.bind("<KeyRelease>", on_key)

    # =========================================================
    # UI
    # =========================================================
    def build_ui(self):
        wrapper = ttk.Frame(self.root)
        wrapper.pack(fill="both", expand=True)

        # TOP BAR
        top = ttk.Frame(wrapper, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Bulan").grid(row=0, column=0)
        self.month_cb = ttk.Combobox(top, values=list(range(1, 13)), width=5)
        self.month_cb.grid(row=0, column=1)
        self.month_cb.set(datetime.now().month)

        ttk.Label(top, text="Tahun").grid(row=0, column=2, padx=(10, 0))
        self.year_cb = ttk.Combobox(top, values=list(range(1970, 2101)), width=7)
        self.year_cb.grid(row=0, column=3)
        self.year_cb.set(datetime.now().year)

        ttk.Button(top, text="Load Bulan", command=self.load_month).grid(row=0, column=4, padx=(10, 0))
        ttk.Button(top, text="Sync Google Sheet", command=self.sync_current_month).grid(row=0, column=5, padx=10)
        ttk.Button(top, text="Export Excel", command=self.export_excel).grid(row=0, column=6)

        # SPLIT
        main = ttk.Frame(wrapper)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # LEFT PANEL
        left = ttk.Frame(main)
        left.pack(side="left", fill="y", padx=(0, 20))

        ttk.Label(left, text="Tanggal").pack(anchor="w")
        self.date_cb = ttk.Combobox(left, width=15)
        self.date_cb.pack(anchor="w")
        self.date_cb.bind("<<ComboboxSelected>>", lambda e: self.load_entry_for_date())

        times = self.make_time_values()

        def add_time(label):
            ttk.Label(left, text=label).pack(anchor="w")
            cb = ttk.Combobox(left, width=15, values=times)
            cb.pack(anchor="w")
            self.enable_autocomplete(cb, times)
            return cb

        self.jm1 = add_time("Jam mulai 1")
        self.js1 = add_time("Jam selesai 1")
        self.jm2 = add_time("Jam mulai 2")
        self.js2 = add_time("Jam selesai 2")
        self.lm = add_time("Lembur mulai")
        self.ls = add_time("Lembur selesai")

        ttk.Label(left, text="Alasan lembur").pack(anchor="w")
        self.alasan = ttk.Combobox(left, values=ALASAN_PRESET, width=30)
        self.alasan.pack(anchor="w")

        ttk.Label(left, text="Deskripsi lembur").pack(anchor="w")
        self.deskripsi = tk.Text(left, width=30, height=4)
        self.deskripsi.pack(anchor="w")

        ttk.Label(left, text="Catatan tambahan").pack(anchor="w")
        self.note = tk.Text(left, width=30, height=4)
        self.note.pack(anchor="w")

        ttk.Button(left, text="Simpan", command=self.save_entry).pack(pady=5)
        ttk.Button(left, text="Hapus", command=self.delete_entry).pack()

        # RIGHT PANEL → TREEVIEW
        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True)

        cols = (
            "Tanggal", "Jam Mulai 1", "Jam Selesai 1",
            "Jam Mulai 2", "Jam Selesai 2", "Total Kerja",
            "Lembur Mulai", "Lembur Selesai", "Total Lembur",
            "Alasan", "Deskripsi", "Catatan"
        )

        self.tree = ttk.Treeview(right, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=130, anchor="center")

        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_tree_double)

        self.status = ttk.Label(self.root, text="Ready")
        self.status.pack(fill="x")

    # =========================================================
    # LOAD BULAN
    # =========================================================
    def load_month(self):
        month = int(self.month_cb.get())
        year = int(self.year_cb.get())

        _, ndays = calendar.monthrange(year, month)
        self.days = [
            date(year, month, i).strftime("%Y-%m-%d")
            for i in range(1, ndays + 1)
        ]

        self.date_cb["values"] = self.days
        self.populate_tree()

    # =========================================================
    # POPULATE TREEVIEW
    # =========================================================
    def populate_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        for d in self.days:
            cur.execute("""
                SELECT jam_mulai_1, jam_selesai_1, jam_mulai_2, jam_selesai_2,
                       lembur_mulai, lembur_selesai,
                       alasan_lembur, deskripsi_lembur, note
                FROM entries WHERE entry_date=?
            """, (d,))
            row = cur.fetchone()

            if row:
                jm1,js1,jm2,js2,lm,ls,alasan,desk,note = row
                tkerja  = calc_total_kerja(jm1,js1,jm2,js2)
                tlembur = calc_total_lembur(lm,ls)
            else:
                jm1=js1=jm2=js2=lm=ls=alasan=desk=note=""
                tkerja = ""
                tlembur = ""

            self.tree.insert(
                "", "end", iid=d,
                values=(d,jm1,js1,jm2,js2,tkerja,lm,ls,tlembur,alasan,desk,note)
            )

        conn.close()

    # =========================================================
    # LOAD ENTRY FORM
    # =========================================================
    def load_entry_for_date(self):
        d = self.date_cb.get()
        if not d:
            return

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        cur.execute("""
            SELECT jam_mulai_1, jam_selesai_1, jam_mulai_2, jam_selesai_2,
                   lembur_mulai, lembur_selesai,
                   alasan_lembur, deskripsi_lembur, note
            FROM entries WHERE entry_date=?
        """, (d,))
        row = cur.fetchone()
        conn.close()

        if row:
            jm1,js1,jm2,js2,lm,ls,alasan,desk,note = row
        else:
            jm1=js1=jm2=js2=lm=ls=alasan=desk=note=""

        self.jm1.set(jm1)
        self.js1.set(js1)
        self.jm2.set(jm2)
        self.js2.set(js2)
        self.lm.set(lm)
        self.ls.set(ls)
        self.alasan.set(alasan)

        self.deskripsi.delete("1.0","end")
        self.deskripsi.insert("1.0", desk)

        self.note.delete("1.0","end")
        self.note.insert("1.0", note)

    # =========================================================
    # SAVE ENTRY
    # =========================================================
    def save_entry(self):
        d = self.date_cb.get()
        if not d:
            messagebox.showerror("Error", "Pilih tanggal dulu")
            return

        jm1 = self.jm1.get()
        js1 = self.js1.get()
        jm2 = self.jm2.get()
        js2 = self.js2.get()
        lm  = self.lm.get()
        ls  = self.ls.get()
        alasan = self.alasan.get()
        desk = self.deskripsi.get("1.0","end").strip()
        note = self.note.get("1.0","end").strip()

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        cur.execute("SELECT id FROM entries WHERE entry_date=?", (d,))
        exists = cur.fetchone()

        if exists:
            cur.execute("""
                UPDATE entries SET
                jam_mulai_1=?, jam_selesai_1=?, jam_mulai_2=?, jam_selesai_2=?,
                lembur_mulai=?, lembur_selesai=?, alasan_lembur=?, deskripsi_lembur=?, note=?
                WHERE entry_date=?
            """, (jm1,js1,jm2,js2,lm,ls,alasan,desk,note,d))
        else:
            cur.execute("""
                INSERT INTO entries (
                    entry_date, jam_mulai_1, jam_selesai_1,
                    jam_mulai_2, jam_selesai_2,
                    lembur_mulai, lembur_selesai,
                    alasan_lembur, deskripsi_lembur, note
                ) VALUES (?,?,?,?,?,?,?,?,?,?)
            """, (d,jm1,js1,jm2,js2,lm,ls,alasan,desk,note))

        conn.commit()
        conn.close()

        self.populate_tree()
        messagebox.showinfo("OK", "Data disimpan.")

    # =========================================================
    # DELETE ENTRY
    # =========================================================
    def delete_entry(self):
        d = self.date_cb.get()
        if not d:
            return

        if not messagebox.askyesno("Hapus?", f"Hapus data {d}?"):
            return

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("DELETE FROM entries WHERE entry_date=?", (d,))
        conn.commit()
        conn.close()

        self.populate_tree()
        self.load_entry_for_date()
        messagebox.showinfo("OK", "Entry dihapus.")

    # =========================================================
    # DOUBLE CLICK → LOAD ENTRY
    # =========================================================
    def on_tree_double(self, event):
        sel = self.tree.focus()
        if sel:
            self.date_cb.set(sel)
            self.load_entry_for_date()

    # =========================================================
    # EXPORT EXCEL EXACT TEMPLATE
    # =========================================================
    def export_excel(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel File",".xlsx")]
        )
        if not path:
            return

        wb = Workbook()
        wb.remove(wb.active)

        month = int(self.month_cb.get())
        year  = int(self.year_cb.get())
        _, ndays = calendar.monthrange(year, month)

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        for day in range(1, ndays+1):
            d = date(year, month, day).strftime("%Y-%m-%d")

            cur.execute("""
                SELECT jam_mulai_1, jam_selesai_1, jam_mulai_2, jam_selesai_2,
                       lembur_mulai, lembur_selesai,
                       alasan_lembur, deskripsi_lembur, note
                FROM entries WHERE entry_date=?
            """, (d,))
            row = cur.fetchone()

            if row:
                jm1,js1,jm2,js2,lm,ls,alasan,desk,note = row
                tkerja = calc_total_kerja(jm1,js1,jm2,js2)
                tlembur = calc_total_lembur(lm,ls)
            else:
                jm1=js1=jm2=js2=lm=ls=alasan=desk=note=""
                tkerja=""
                tlembur=""

            ws = wb.create_sheet(d)
            template = [
                ["Identitas"],
                ["Nama lengkap","Refia Karsista"],
                [],
                ["Tanggal lembur", d],
                [],
                ["Waktu Kerja"],
                ["Jam mulai 1", jm1],
                ["Jam selesai 1", js1],
                ["Jam mulai 2", jm2],
                ["Jam selesai 2", js2],
                ["Total Waktu Kerja", tkerja],
                [],
                ["Tanggal & Waktu Lembur"],
                ["Jam mulai lembur", lm],
                ["Jam selesai lembur", ls],
                ["Jam mulai lembur 2", "-"],
                ["Jam selesai lembur 2", "-"],
                ["Total Lembur", tlembur],
                [],
                ["Alasan Lembur", alasan],
                [],
                ["Deskripsi Pekerjaan", desk],
                [],
                ["Catatan Tambahan", note],
            ]

            for rowdata in template:
                ws.append(rowdata)

        conn.close()
        wb.save(path)
        messagebox.showinfo("OK", "Export Excel selesai.")

    # =========================================================
    # SYNC GOOGLE SHEET FINAL
    # =========================================================
    def sync_current_month(self):
        try:
            self.status.config(text="Syncing Google Sheet...")
            self.root.update_idletasks()

            if not self.gs:
                self.gs = GoogleSheetSync(
                    resource_path("credentials.json"),
                    "AI META Timesheet"
                )

            month = int(self.month_cb.get())
            year  = int(self.year_cb.get())
            _, ndays = calendar.monthrange(year, month)

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            for day in range(1, ndays + 1):
                d = date(year, month, day).strftime("%Y-%m-%d")

                cur.execute("""
                    SELECT jam_mulai_1, jam_selesai_1,
                        jam_mulai_2, jam_selesai_2,
                        lembur_mulai, lembur_selesai,
                        alasan_lembur, deskripsi_lembur, note
                    FROM entries
                    WHERE entry_date=?
                """, (d,))
                row = cur.fetchone()

                if row:
                    jm1, js1, jm2, js2, lm, ls, alasan, desk, note = row
                    tkerja  = calc_total_kerja(jm1, js1, jm2, js2)
                    tlembur = calc_total_lembur(lm, ls)
                else:
                    jm1 = js1 = jm2 = js2 = lm = ls = alasan = desk = note = ""
                    tkerja = ""
                    tlembur = ""

                data = {
                    "nama": "Refia Karsista",
                    "jam_mulai_1": jm1,
                    "jam_selesai_1": js1,
                    "jam_mulai_2": jm2,
                    "jam_selesai_2": js2,
                    "total_kerja": tkerja,
                    "lembur_mulai": lm,
                    "lembur_selesai": ls,
                    "total_lembur": tlembur,
                    "alasan_lembur": alasan,
                    "deskripsi_lembur": desk,
                    "catatan": note,
                }

                self.status.config(text=f"Sync {d}...")
                self.root.update_idletasks()

                self.gs.write_daily_sheet(d, data)
                time.sleep(1)  # to avoid rate limit

            conn.close()
            messagebox.showinfo(
                "OK",
                f"Sync Google Sheet berhasil ({month}/{year})"
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("ERROR", str(e))

        finally:
            self.status.config(text="Ready")


# =============================================================
# RUN PROGRAM
# =============================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = TimesheetApp(root)
    root.mainloop()
