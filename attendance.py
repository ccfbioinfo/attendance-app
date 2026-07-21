"""
Staff Duty Attendance System v3.0
- No manual edit of attendance records
- Smart auto-completion in monthly report:
  - Partial checkin/checkout: fill missing time from schedule
  - No records: mark as "Forgot Check" or "Off Day" based on schedule
- 8.8 hrs threshold with 5-min grace
"""

import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import os
import sys
import csv
import re

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ---------- Configuration ----------
WORK_START = "09:00:00"
WORK_END   = "18:00:00"
CURRENT_YEAR = datetime.date.today().year
STANDARD_HOURS = 8.8
GRACE_MINUTES = 5
GRACE_HOURS = GRACE_MINUTES / 60.0

# ---------- Shift Code Mapping ----------
SHIFT_MAP = {
    "MD":   ("08:00", "16:48"),
    "D8":   ("08:00", "16:48"),
    "R8":   ("08:00", "16:48"),
    "D/SD": ("09:00", "17:48"),
    "R10":  ("10:00", "18:48"),
    "P":    ("13:00", "21:48"),
    "N":    ("21:30", "08:30"),
    "W":    ("08:30", "17:18"),
    "AA1":  ("09:00", "18:00"),
    "AA2":  ("09:00", "17:00"),
    "D4":   ("08:00", "16:48"),
    "D1":   ("08:00", "16:48"),
    "D3":   ("08:00", "16:48"),
    "D5":   ("08:00", "16:48"),
    "D6":   ("08:00", "16:48"),
    "MD1":  ("08:00", "16:48"),
    "MD2":  ("08:00", "16:48"),
    "MD3":  ("08:00", "16:48"),
    "MD4":  ("08:00", "16:48"),
    "MD5":  ("08:00", "16:48"),
    "SD":   ("09:00", "17:48"),
    "PH":   ("09:00", "17:48"),
    "Ag/gP": ("13:00", "21:48"),
    "Ag2/gP2": ("13:00", "21:48"),
    "O":    None,
    "AL":   None,
    "SL":   None,
    "AM AL/ PM SL": None,
    "D/ PM NPL": None,
    "AA2/ PM SL": None,
}

# ---------- Database Path: Always next to .exe ----------
def get_db_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, 'attendance.db')

DB_PATH = get_db_path()

# ---------- Database Setup ----------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS staff (
            staff_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            batch TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            staff_id TEXT NOT NULL,
            date TEXT NOT NULL,
            checkin TEXT,
            checkout TEXT,
            FOREIGN KEY (staff_id) REFERENCES staff(staff_id),
            UNIQUE(staff_id, date)
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS work_schedule (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            staff_id TEXT NOT NULL,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            work_start TEXT NOT NULL,
            work_end TEXT NOT NULL,
            FOREIGN KEY (staff_id) REFERENCES staff(staff_id),
            UNIQUE(staff_id, start_date, end_date)
        )
    ''')
    conn.commit()
    conn.close()

# ---------- Database Helpers ----------
def get_staff(staff_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT staff_id, name, batch FROM staff WHERE staff_id=?", (staff_id,))
    row = c.fetchone()
    conn.close()
    return row

def get_staff_by_name(name):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT staff_id FROM staff WHERE name=?", (name,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else None

def get_today_attendance(staff_id):
    today = datetime.date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT checkin, checkout FROM attendance WHERE staff_id=? AND date=?", (staff_id, today))
    row = c.fetchone()
    conn.close()
    return row

def set_checkin(staff_id, time_str):
    today = datetime.date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO attendance (staff_id, date, checkin, checkout)
        VALUES (?, ?, ?, NULL)
    ''', (staff_id, today, time_str))
    conn.commit()
    conn.close()

def set_checkout(staff_id, time_str):
    today = datetime.date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        UPDATE attendance SET checkout=?
        WHERE staff_id=? AND date=? AND checkin IS NOT NULL
    ''', (time_str, staff_id, today))
    conn.commit()
    conn.close()

def override_checkin(staff_id, time_str):
    today = datetime.date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        UPDATE attendance SET checkin=?, checkout=NULL
        WHERE staff_id=? AND date=?
    ''', (time_str, staff_id, today))
    conn.commit()
    conn.close()

def upsert_attendance(staff_id, date_str, checkin, checkout):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO attendance (staff_id, date, checkin, checkout)
        VALUES (?, ?, ?, ?)
    ''', (staff_id, date_str, checkin, checkout))
    conn.commit()
    conn.close()

def upsert_work_schedule(staff_id, start_date, end_date, work_start, work_end):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO work_schedule (staff_id, start_date, end_date, work_start, work_end)
        VALUES (?, ?, ?, ?, ?)
    ''', (staff_id, start_date, end_date, work_start, work_end))
    conn.commit()
    conn.close()

def get_work_schedule_for_date(staff_id, date_str):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT work_start, work_end
        FROM work_schedule
        WHERE staff_id = ? AND ? BETWEEN start_date AND end_date
        ORDER BY start_date DESC LIMIT 1
    ''', (staff_id, date_str))
    row = c.fetchone()
    conn.close()
    return row if row else None

def get_monthly_attendance(year, month):
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year+1}-01-01"
    else:
        end_date = f"{year}-{month+1:02d}-01"
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT s.staff_id, s.name, s.batch, a.date, a.checkin, a.checkout
        FROM attendance a
        JOIN staff s ON a.staff_id = s.staff_id
        WHERE a.date >= ? AND a.date < ?
          AND a.checkin IS NOT NULL AND a.checkout IS NOT NULL
        ORDER BY s.staff_id, a.date
    ''', (start_date, end_date))
    rows = c.fetchall()
    conn.close()
    return rows

def calculate_work_hours(checkin_str, checkout_str):
    try:
        ci = datetime.datetime.strptime(checkin_str, "%H:%M:%S")
        co = datetime.datetime.strptime(checkout_str, "%H:%M:%S")
        if co <= ci:
            co += datetime.timedelta(days=1)
        delta = co - ci
        return delta.total_seconds() / 3600.0
    except:
        return 0.0

# ---------- GUI Application ----------
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Staff Attendance System v3.0")
        self.root.geometry("700x550")
        self.show_db_path()
        self.confirm_dialog = None
        self.timer_id = None
        self.countdown = 10
        self.current_staff_id = None
        self.current_name = None
        self.current_batch = None
        self.create_widgets()
        self.barcode_entry.bind("<Return>", self.on_barcode_scan)
        self.update_status()

    def show_db_path(self):
        messagebox.showinfo("Database Location",
                            f"Attendance records stored at:\n{DB_PATH}\n\n"
                            f"Standard work hours: {STANDARD_HOURS} hrs (±{GRACE_MINUTES} min grace)\n"
                            "Missing check-in/out will be auto-completed in reports.")

    def create_widgets(self):
        top_frame = ttk.LabelFrame(self.root, text="Scan Barcode", padding=10)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(top_frame, text="Scan / Enter Staff ID:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.barcode_entry = ttk.Entry(top_frame, width=30)
        self.barcode_entry.grid(row=0, column=1, padx=5, pady=5)
        self.barcode_entry.focus_set()

        self.scan_btn = ttk.Button(top_frame, text="Process Scan", command=self.on_barcode_scan)
        self.scan_btn.grid(row=0, column=2, padx=5, pady=5)

        info_frame = ttk.LabelFrame(self.root, text="Staff Info", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=5)

        self.name_var = tk.StringVar(value="")
        self.batch_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Ready")

        ttk.Label(info_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Label(info_frame, textvariable=self.name_var).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text="Batch:").grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(info_frame, textvariable=self.batch_var).grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text="Status:").grid(row=2, column=0, sticky=tk.W, padx=5)
        ttk.Label(info_frame, textvariable=self.status_var).grid(row=2, column=1, sticky=tk.W, padx=5)

        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(btn_frame, text="Add / Edit Staff", command=self.manage_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Import Attendance", command=self.import_attendance).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Import Roster (Excel)", command=self.import_roster).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Monthly Exceptions", command=self.show_monthly_summary).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Full Monthly Report", command=self.export_full_monthly_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

        log_frame = ttk.LabelFrame(self.root, text="Recent Activity", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log_message(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%H:%M:%S')} - {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def update_status(self, staff_id=None):
        if staff_id:
            self.current_staff_id = staff_id
            staff = get_staff(staff_id)
            if staff:
                self.current_name = staff[1]
                self.current_batch = staff[2]
                self.name_var.set(self.current_name)
                self.batch_var.set(self.current_batch)
                att = get_today_attendance(staff_id)
                if att:
                    checkin_time = att[0]
                    checkout_time = att[1]
                    if checkout_time:
                        self.status_var.set(f"Checked out at {checkout_time}")
                    else:
                        self.status_var.set(f"Checked in at {checkin_time}")
                else:
                    self.status_var.set("Not checked in today")
            else:
                self.current_name = None
                self.current_batch = None
                self.name_var.set("Unknown")
                self.batch_var.set("")
                self.status_var.set("Staff ID not found")
        else:
            self.current_staff_id = None
            self.current_name = None
            self.current_batch = None
            self.name_var.set("")
            self.batch_var.set("")
            self.status_var.set("Ready")

    # ---------- Import Attendance (CSV/Excel) ----------
    def import_attendance(self):
        file_path = filedialog.askopenfilename(
            title="Select attendance file",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            self.import_attendance_csv(file_path)
        elif ext in ('.xlsx', '.xls'):
            self.import_attendance_excel(file_path)
        else:
            messagebox.showerror("Error", "Unsupported format.")

    def import_attendance_csv(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader, None)
                count = 0
                for row in reader:
                    if len(row) < 4:
                        continue
                    staff_id, date_str, checkin, checkout = row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip()
                    try:
                        datetime.datetime.strptime(date_str, "%Y-%m-%d")
                    except ValueError:
                        self.log_message(f"Skipping invalid date: {date_str}")
                        continue
                    if not get_staff(staff_id):
                        self.log_message(f"Staff {staff_id} not found, skipping.")
                        continue
                    upsert_attendance(staff_id, date_str, checkin, checkout)
                    count += 1
                self.log_message(f"Imported {count} attendance records.")
                messagebox.showinfo("Success", f"Imported {count} records.")
        except Exception as e:
            messagebox.showerror("Error", f"CSV import failed: {str(e)}")

    def import_attendance_excel(self, file_path):
        if not HAS_OPENPYXL:
            messagebox.showerror("Error", "openpyxl not installed. Please install: pip install openpyxl")
            return
        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and len(row) >= 4:
                    staff_id = str(row[0]).strip()
                    date_str = str(row[1]).strip()
                    checkin = str(row[2]).strip() if row[2] else None
                    checkout = str(row[3]).strip() if row[3] else None
                    if not staff_id or not date_str:
                        continue
                    try:
                        datetime.datetime.strptime(date_str, "%Y-%m-%d")
                    except ValueError:
                        self.log_message(f"Skipping invalid date: {date_str}")
                        continue
                    if not get_staff(staff_id):
                        self.log_message(f"Staff {staff_id} not found, skipping.")
                        continue
                    upsert_attendance(staff_id, date_str, checkin, checkout)
                    count += 1
            self.log_message(f"Imported {count} attendance records from Excel.")
            messagebox.showinfo("Success", f"Imported {count} records.")
        except Exception as e:
            messagebox.showerror("Error", f"Excel import failed: {str(e)}")

    # ---------- Import Roster ----------
    def import_roster(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("Error", "openpyxl is required for roster import. Please install: pip install openpyxl")
            return

        file_path = filedialog.askopenfilename(
            title="Select roster Excel file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active

            date_row_idx = None
            date_cols = []
            for row_idx in range(1, min(15, ws.max_row + 1)):
                row_values = [cell.value for cell in ws[row_idx]]
                found = False
                for col_idx, val in enumerate(row_values, start=1):
                    if val and isinstance(val, str):
                        parsed = self.parse_date(val)
                        if parsed:
                            date_cols.append((col_idx, parsed))
                            found = True
                if found:
                    date_row_idx = row_idx
                    break

            if not date_cols:
                messagebox.showerror("Error", "No date columns found in the first 15 rows.")
                return

            date_cols.sort(key=lambda x: x[0])
            first_date_col = date_cols[0][0]
            name_col = first_date_col - 1
            if name_col < 1:
                messagebox.showerror("Error", "Name column not found (must be before date columns).")
                return

            self.log_message(f"Detected date row: {date_row_idx}, name column: {name_col}, first date col: {first_date_col}")

            inserted = 0
            skipped = 0
            for row_idx in range(date_row_idx + 1, ws.max_row + 1):
                name_cell = ws.cell(row=row_idx, column=name_col)
                name = str(name_cell.value).strip() if name_cell.value else None
                if not name:
                    continue

                staff_id = get_staff_by_name(name)
                if not staff_id:
                    self.log_message(f"Staff '{name}' not found in database, skipping row {row_idx}")
                    skipped += 1
                    continue

                for col_idx, date_str in date_cols:
                    cell = ws.cell(row=row_idx, column=col_idx)
                    shift_code = str(cell.value).strip() if cell.value else ""
                    if not shift_code:
                        continue
                    if shift_code in SHIFT_MAP:
                        times = SHIFT_MAP[shift_code]
                        if times is None:
                            continue
                        work_start, work_end = times
                        upsert_work_schedule(staff_id, date_str, date_str, work_start, work_end)
                        inserted += 1
                    else:
                        self.log_message(f"Unknown shift code '{shift_code}' for {name} on {date_str}")

            self.log_message(f"Roster import completed. Inserted {inserted} schedules, skipped {skipped} staff rows.")
            messagebox.showinfo("Import Complete", f"Imported {inserted} work schedules from roster.\nSkipped {skipped} rows due to missing staff.")

        except Exception as e:
            messagebox.showerror("Error", f"Roster import failed: {str(e)}")

    def parse_date(self, date_str):
        s = date_str.strip()
        match = re.match(r'^(\d{1,2})-(\d{1,2})月$', s)
        if match:
            day = int(match.group(1))
            month = int(match.group(2))
            if 1 <= month <= 12 and 1 <= day <= 31:
                try:
                    dt = datetime.date(CURRENT_YEAR, month, day)
                    return dt.isoformat()
                except ValueError:
                    return None
        match = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{2})$', s)
        if match:
            day = int(match.group(1))
            month = int(match.group(2))
            year_short = int(match.group(3))
            year = 2000 + year_short if year_short < 100 else year_short
            if 1 <= month <= 12 and 1 <= day <= 31:
                try:
                    dt = datetime.date(year, month, day)
                    return dt.isoformat()
                except ValueError:
                    return None
        return None

    # ---------- Barcode Scan ----------
    def on_barcode_scan(self, event=None):
        if self.confirm_dialog is not None and self.confirm_dialog.winfo_exists():
            self.log_message("Scan ignored – confirmation pending")
            self.barcode_entry.delete(0, tk.END)
            return

        barcode = self.barcode_entry.get().strip()
        if not barcode:
            return
        self.barcode_entry.delete(0, tk.END)

        staff = get_staff(barcode)
        if not staff:
            if messagebox.askyesno("Staff Not Found", f"Staff ID '{barcode}' not found.\nDo you want to add this staff now?"):
                self.add_new_staff(barcode)
                self.update_status(barcode)
            else:
                self.log_message(f"Unknown barcode: {barcode}")
            return

        staff_id, name, batch = staff
        self.log_message(f"Scanned: {name} ({staff_id})")
        self.update_status(staff_id)

        att = get_today_attendance(staff_id)
        now = datetime.datetime.now().strftime("%H:%M:%S")
        if att:
            checkin_time = att[0]
            checkout_time = att[1]
            if checkout_time:
                action = "Check-in (override)"
                action_key = "override"
            else:
                action = "Check-out"
                action_key = "checkout"
        else:
            action = "Check-in"
            action_key = "checkin"

        self.show_confirmation(staff_id, name, batch, action, action_key, now)

    def show_confirmation(self, staff_id, name, batch, action, action_key, current_time):
        dialog = tk.Toplevel(self.root)
        dialog.title("Confirm Attendance")
        dialog.geometry("400x280")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()

        self.barcode_entry.config(state=tk.DISABLED)
        self.scan_btn.config(state=tk.DISABLED)

        self.confirm_dialog = dialog
        self.countdown = 10

        ttk.Label(dialog, text="Staff:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(dialog, text=f"{name} ({staff_id})", font=("Arial", 12, "bold")).grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

        ttk.Label(dialog, text="Batch:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(dialog, text=batch or "-", font=("Arial", 12)).grid(row=1, column=1, padx=10, pady=5, sticky=tk.W)

        ttk.Label(dialog, text="Current time:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(dialog, text=current_time, font=("Arial", 12, "bold")).grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)

        ttk.Label(dialog, text="Action:", font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Label(dialog, text=action, font=("Arial", 12, "bold"), foreground="blue").grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)

        self.countdown_label = ttk.Label(dialog, text=f"Auto‑confirm in {self.countdown} seconds", font=("Arial", 10))
        self.countdown_label.grid(row=4, column=0, columnspan=2, pady=10)

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=15)

        def reset_after_dialog():
            self.barcode_entry.config(state=tk.NORMAL)
            self.scan_btn.config(state=tk.NORMAL)
            self.barcode_entry.focus_set()
            self.confirm_dialog = None
            if self.timer_id:
                self.root.after_cancel(self.timer_id)
                self.timer_id = None

        def do_confirm():
            if self.timer_id:
                dialog.after_cancel(self.timer_id)
                self.timer_id = None
            self.perform_action(staff_id, action_key, current_time)
            reset_after_dialog()
            dialog.destroy()

        def do_cancel():
            if self.timer_id:
                dialog.after_cancel(self.timer_id)
                self.timer_id = None
            self.log_message(f"Confirmation cancelled for {name}")
            reset_after_dialog()
            dialog.destroy()

        confirm_btn = ttk.Button(btn_frame, text="Confirm", command=do_confirm, width=12)
        confirm_btn.pack(side=tk.LEFT, padx=10)
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=do_cancel, width=12)
        cancel_btn.pack(side=tk.LEFT, padx=10)

        self.update_countdown(dialog, staff_id, action_key, current_time, reset_after_dialog)
        dialog.protocol("WM_DELETE_WINDOW", do_cancel)

    def update_countdown(self, dialog, staff_id, action_key, current_time, reset_callback):
        if self.countdown <= 0:
            if self.timer_id:
                self.timer_id = None
            self.perform_action(staff_id, action_key, current_time)
            reset_callback()
            dialog.destroy()
            return

        self.countdown_label.config(text=f"Auto‑confirm in {self.countdown} seconds")
        self.countdown -= 1
        self.timer_id = dialog.after(1000, self.update_countdown, dialog, staff_id, action_key, current_time, reset_callback)

    def perform_action(self, staff_id, action_key, time_str):
        if action_key == "checkin":
            set_checkin(staff_id, time_str)
            self.log_message(f"Checked in at {time_str}")
        elif action_key == "checkout":
            set_checkout(staff_id, time_str)
            self.log_message(f"Checked out at {time_str}")
        elif action_key == "override":
            override_checkin(staff_id, time_str)
            self.log_message(f"Overrode check‑in at {time_str} (previous checkout cleared)")
        else:
            self.log_message("Unknown action – nothing stored")
            return
        self.update_status(staff_id)

    # ---------- Staff Management (No Edit Attendance) ----------
    def manage_staff(self):
        win = tk.Toplevel(self.root)
        win.title("Manage Staff")
        win.geometry("500x400")

        tree = ttk.Treeview(win, columns=("ID", "Name", "Batch"), show="headings")
        tree.heading("ID", text="Staff ID")
        tree.heading("Name", text="Name")
        tree.heading("Batch", text="Batch")
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        def refresh_staff_list():
            for row in tree.get_children():
                tree.delete(row)
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
            for r in c.fetchall():
                tree.insert("", tk.END, values=r)
            conn.close()

        refresh_staff_list()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=5)

        def add_staff():
            staff_id = simpledialog.askstring("Add Staff", "Enter staff ID:", parent=win)
            if staff_id:
                if get_staff(staff_id):
                    messagebox.showerror("Error", "Staff ID already exists.")
                    return
                name = simpledialog.askstring("Add Staff", "Enter name:", parent=win)
                if name:
                    batch = simpledialog.askstring("Add Staff", "Enter batch:", parent=win)
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    c.execute("INSERT INTO staff (staff_id, name, batch) VALUES (?, ?, ?)",
                              (staff_id, name, batch))
                    conn.commit()
                    conn.close()
                    refresh_staff_list()

        def edit_staff():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("Info", "Select a staff to edit.")
                return
            values = tree.item(selected[0])['values']
            staff_id = values[0]
            name = values[1]
            batch = values[2]
            new_name = simpledialog.askstring("Edit Staff", "New name:", initialvalue=name, parent=win)
            if new_name:
                new_batch = simpledialog.askstring("Edit Staff", "New batch:", initialvalue=batch, parent=win)
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("UPDATE staff SET name=?, batch=? WHERE staff_id=?", (new_name, new_batch, staff_id))
                conn.commit()
                conn.close()
                refresh_staff_list()

        def delete_staff():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("Info", "Select a staff to delete.")
                return
            values = tree.item(selected[0])['values']
            staff_id = values[0]
            if messagebox.askyesno("Delete", f"Delete staff '{values[1]}' and all attendance records?"):
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("DELETE FROM attendance WHERE staff_id=?", (staff_id,))
                c.execute("DELETE FROM staff WHERE staff_id=?", (staff_id,))
                conn.commit()
                conn.close()
                refresh_staff_list()

        ttk.Button(btn_frame, text="Add", command=add_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Edit", command=edit_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", command=delete_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side=tk.RIGHT, padx=5)

    # ---------- Add New Staff (helper) ----------
    def add_new_staff(self, staff_id):
        name = simpledialog.askstring("Add Staff", "Enter staff name:", parent=self.root)
        if name:
            batch = simpledialog.askstring("Add Staff", "Enter batch (optional):", parent=self.root)
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            try:
                c.execute("INSERT INTO staff (staff_id, name, batch) VALUES (?, ?, ?)",
                          (staff_id, name, batch))
                conn.commit()
                self.log_message(f"Added staff: {name} ({staff_id})")
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Staff ID already exists.")
            conn.close()

    # ---------- Monthly Exception Report ----------
    def show_monthly_summary(self):
        win = tk.Toplevel(self.root)
        win.title("Monthly Exception Report")
        win.geometry("400x200")

        ttk.Label(win, text="Select Month:").pack(pady=10)
        ttk.Label(win, text="Enter year and month (YYYY-MM):").pack()
        entry = ttk.Entry(win)
        entry.pack(pady=5)
        entry.insert(0, datetime.date.today().strftime("%Y-%m"))

        def generate():
            date_str = entry.get()
            try:
                year, month = map(int, date_str.split('-'))
            except:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM")
                return

            all_records = get_monthly_attendance(year, month)
            if not all_records:
                messagebox.showinfo("No Data", "No attendance records for this month.")
                return

            exceptions = {}
            for row in all_records:
                staff_id, name, batch, date_str, checkin_str, checkout_str = row
                schedule = get_work_schedule_for_date(staff_id, date_str)
                if schedule:
                    work_start_str, work_end_str = schedule
                else:
                    work_start_str, work_end_str = WORK_START, WORK_END
                try:
                    checkin_time = datetime.datetime.strptime(checkin_str, "%H:%M:%S").time()
                    checkout_time = datetime.datetime.strptime(checkout_str, "%H:%M:%S").time()
                    work_start = datetime.datetime.strptime(work_start_str, "%H:%M:%S").time()
                    work_end = datetime.datetime.strptime(work_end_str, "%H:%M:%S").time()
                except ValueError:
                    continue

                issues = []
                if checkin_time > work_start:
                    issues.append(("Late", checkin_str))
                if checkout_time < work_end:
                    issues.append(("Early Leave", checkout_str))

                if issues:
                    if staff_id not in exceptions:
                        exceptions[staff_id] = {'name': name, 'batch': batch, 'days': []}
                    for issue_type, time_val in issues:
                        exceptions[staff_id]['days'].append((date_str, issue_type, time_val))

            if not exceptions:
                messagebox.showinfo("All Good", "No late arrivals or early departures for this month.")
                return

            result_win = tk.Toplevel(win)
            result_win.title(f"Exception Report - {year}-{month:02d}")
            result_win.geometry("800x500")

            tree = ttk.Treeview(result_win, columns=("Staff", "Batch", "Date", "Issue", "Time"), show="headings")
            tree.heading("Staff", text="Staff (ID)")
            tree.heading("Batch", text="Batch")
            tree.heading("Date", text="Date")
            tree.heading("Issue", text="Issue Type")
            tree.heading("Time", text="Recorded Time")
            tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            for staff_id, data in exceptions.items():
                for date_str, issue_type, time_val in data['days']:
                    tree.insert("", tk.END, values=(
                        f"{data['name']} ({staff_id})",
                        data['batch'] or "-",
                        date_str,
                        issue_type,
                        time_val
                    ))

            def export_exceptions():
                import csv
                file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                         filetypes=[("CSV files", "*.csv")])
                if not file_path:
                    return
                with open(file_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Staff ID", "Name", "Batch", "Date", "Issue Type", "Recorded Time"])
                    for staff_id, data in exceptions.items():
                        for date_str, issue_type, time_val in data['days']:
                            writer.writerow([staff_id, data['name'], data['batch'] or "", date_str, issue_type, time_val])
                messagebox.showinfo("Export", f"Report exported to {file_path}")

            ttk.Button(result_win, text="Export CSV", command=export_exceptions).pack(pady=5)

        ttk.Button(win, text="Generate", command=generate).pack(pady=20)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=5)

    # ---------- Full Monthly Report with Auto-Completion ----------
    def export_full_monthly_report(self):
        win = tk.Toplevel(self.root)
        win.title("Full Monthly Report")
        win.geometry("400x200")

        ttk.Label(win, text="Select Month:").pack(pady=10)
        ttk.Label(win, text="Enter year and month (YYYY-MM):").pack()
        entry = ttk.Entry(win)
        entry.pack(pady=5)
        entry.insert(0, datetime.date.today().strftime("%Y-%m"))

        def generate():
            date_str = entry.get()
            try:
                year, month = map(int, date_str.split('-'))
            except:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM")
                return

            # Get all staff
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
            all_staff = c.fetchall()
            conn.close()

            if not all_staff:
                messagebox.showinfo("No Staff", "No staff records found.")
                return

            # Generate list of all days in the month
            first_day = datetime.date(year, month, 1)
            if month == 12:
                next_month = datetime.date(year + 1, 1, 1)
            else:
                next_month = datetime.date(year, month + 1, 1)
            all_dates = []
            current = first_day
            while current < next_month:
                all_dates.append(current.isoformat())
                current += datetime.timedelta(days=1)

            report_data = []

            for staff_id, name, batch in all_staff:
                # Get all attendance records for this staff in the month
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute('''
                    SELECT date, checkin, checkout
                    FROM attendance
                    WHERE staff_id=? AND date >= ? AND date < ?
                ''', (staff_id, first_day.isoformat(), next_month.isoformat()))
                records = {row[0]: (row[1], row[2]) for row in c.fetchall()}
                conn.close()

                for date_str in all_dates:
                    # Check schedule for this day
                    schedule = get_work_schedule_for_date(staff_id, date_str)
                    if not schedule:
                        # No schedule => Off Day
                        report_data.append({
                            "Staff ID": staff_id,
                            "Name": name,
                            "Batch": batch or "",
                            "Date": date_str,
                            "Checkin": "",
                            "Checkout": "",
                            "Work Hours": "0.00",
                            "Status": "Off Day"
                        })
                        continue

                    work_start, work_end = schedule

                    # Check actual attendance
                    if date_str in records:
                        checkin, checkout = records[date_str]
                        # If only checkin exists, fill checkout
                        if checkin and not checkout:
                            checkout = work_end
                            status_note = " (Auto-completed checkout)"
                        elif not checkin and checkout:
                            checkin = work_start
                            status_note = " (Auto-completed checkin)"
                        elif checkin and checkout:
                            status_note = ""
                        else:
                            # Should not happen, but if both None, treat as forgot
                            checkin = work_start
                            checkout = work_end
                            status_note = " (No data, auto-filled)"
                    else:
                        # No attendance record at all => Forgot Check
                        checkin = work_start
                        checkout = work_end
                        status_note = " (Forgot Check)"

                    # Calculate hours and determine status
                    work_hrs = calculate_work_hours(checkin, checkout) if checkin and checkout else 0.0
                    if work_hrs == 0:
                        status = "No Data"
                    elif abs(work_hrs - STANDARD_HOURS) <= GRACE_HOURS:
                        status = "Normal"
                    elif work_hrs < STANDARD_HOURS - GRACE_HOURS:
                        status = "Early Leave"
                    else:
                        status = "Overtime"

                    # Append status note if auto-completed or forgot
                    if status_note:
                        status += status_note

                    report_data.append({
                        "Staff ID": staff_id,
                        "Name": name,
                        "Batch": batch or "",
                        "Date": date_str,
                        "Checkin": checkin or "",
                        "Checkout": checkout or "",
                        "Work Hours": f"{work_hrs:.2f}",
                        "Status": status
                    })

            if not report_data:
                messagebox.showinfo("No Data", "No data to export.")
                return

            # Save CSV
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Save Full Monthly Report"
            )
            if not file_path:
                return

            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=["Staff ID", "Name", "Batch", "Date", "Checkin", "Checkout", "Work Hours", "Status"])
                writer.writeheader()
                writer.writerows(report_data)

            messagebox.showinfo("Export Complete", f"Full monthly report saved to:\n{file_path}")

        ttk.Button(win, text="Generate & Export CSV", command=generate).pack(pady=20)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=5)


# ---------- Main ----------
if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()