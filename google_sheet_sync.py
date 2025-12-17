import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime


class GoogleSheetSync:
    def __init__(self, creds_path: str, sheet_name: str):
        scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
        client = gspread.authorize(creds)
        self.sheet = client.open(sheet_name)


def ensure_daily_sheet(self, date_str: str):
    try:
        ws = self.sheet.worksheet(date_str)
    except gspread.exceptions.WorksheetNotFound:
        ws = self.sheet.add_worksheet(title=date_str, rows="200", cols="10")
    return ws

def write_daily_sheet(self, date_str: str, data: dict):
    ws = self.ensure_daily_sheet(date_str)
    rows = [
        ["Identitas"],
        ["Nama lengkap", data.get("nama", "Dummy")],
        [],
        ["Tanggal lembur", date_str],
        [],
        ["Waktu Kerja Work Day"],
        ["Jam mulai 1", data.get("jam_mulai_1", "")],
        ["Jam selesai 1", data.get("jam_selesai_1", "")],
        ["Jam mulai 2", data.get("jam_mulai_2", "")],
        ["Jam selesai 2", data.get("jam_selesai_2", "")],
        ["Total Waktu Kerja", data.get("total_kerja", "")],
        [],
        ["Tanggal & Waktu Lembur"],
        ["Jam mulai lembur", data.get("lembur_mulai", "")],
        ["Jam selesai lembur", data.get("lembur_selesai", "")],
        ["Jam mulai lembur 2", "-"] ,
        ["Jam selesai lembur 2", "-"],
        ["Total Lembur", data.get("total_lembur", "")],
        [],
        ["Alasan / Kebutuhan Lembur", data.get("alasan_lembur","")],
        [],
        ["Deskripsi Pekerjaan yang Dikerjakan", data.get("deskripsi_lembur","")],
        [],
        ["Catatan Tambahan", data.get("catatan", "")],
    ]
    ws.clear()
    for row in rows:
        ws.append_row(row)