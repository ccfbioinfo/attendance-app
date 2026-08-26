"""
Staff Duty Attendance System v5.7
- Removed Import Roster and Shift Code Reference buttons.
- Added shift_code column to attendance table.
- Clock-in window shows shift selection; Clock-out window shows only leave reasons.
- Updated terminology: check-in/out -> clock-in/out.
"""

import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import os
import sys
import csv
import re
import threading
import shutil
import traceback

# Note: openpyxl is no longer required (roster import removed), but keep import for potential future use.
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
GRACE_MINUTES = 0
GRACE_HOURS = GRACE_MINUTES / 60.0

# ---------- Shift Code Mapping ----------
SHIFT_ENTRIES = [
    ("B", "07:45", "16:33"),
    ("OA", "07:45", "16:33"),
    ("MD", "08:00", "16:48"),
    ("D8", "08:00", "16:48"),
    ("R8", "08:00", "16:48"),
    ("SAT_LA", "08:00", "16:48"),
    ("A", "08:15", "17:03"),
    ("C", "08:15", "17:03"),
    ("SAT_Tech", "08:15", "17:03"),
    ("D_AP", "08:15", "17:03"),
    ("W", "08:30", "17:18"),
    ("S", "08:45", "17:33"),
    ("D", "09:00", "17:48"),
    ("SD", "09:00", "17:48"),
    ("AA1", "09:00", "18:00"),
    ("AA2", "09:00", "17:00"),
    ("P_AP", "09:15", "18:03"),
    ("OB*", "09:15", "18:03"),
    ("OB", "09:45", "18:33"),
    ("R10", "10:00", "18:48"),
    ("N_AP", "10:15", "19:03"),
    ("P_Core", "13:00", "21:48"),
    ("N_Core", "21:30", "08:30"),
]

SHIFT_MAP = {}
for code, start, end in SHIFT_ENTRIES:
    SHIFT_MAP[code] = (start, end)

REST_CODES = ["O", "AL", "SL", "AM AL/ PM SL", "D/ PM NPL", "AA2/ PM SL"]
for code in REST_CODES:
    SHIFT_MAP[code] = None

# ---------- Helper Functions ----------
def pad_time(t):
    if not t:
        return t
    t = t.strip()
    if len(t) == 5 and ':' in t:
        return t + ":00"
    return t

def get_db_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, 'attendance.db')

DB_PATH = get_db_path()

def get_backup_dir():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    backup_dir = os.path.join(base_dir, 'backup')
    try:
        os.makedirs(backup_dir, exist_ok=True)
        test_file = os.path.join(backup_dir, 'test.tmp')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        return backup_dir
    except:
        if os.name == 'nt':
            appdata = os.getenv('APPDATA')
            if not appdata:
                appdata = os.path.expanduser('~/AppData/Roaming')
            backup_dir = os.path.join(appdata, 'AttendanceSystem', 'backup')
        else:
            backup_dir = os.path.expanduser('~/.local/share/AttendanceSystem/backup')
        os.makedirs(backup_dir, exist_ok=True)
        return backup_dir

def get_today_log_path():
    backup_dir = get_backup_dir()
    today = datetime.date.today().isoformat()
    return os.path.join(backup_dir, f"{today}.log")

def write_log_to_file(msg):
    try:
        log_path = get_today_log_path()
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp} - {msg}\n")
    except Exception as e:
        print(f"Failed to write log: {e}")

# ---------- Database Backup ----------
def backup_db():
    try:
        src = DB_PATH
        if not os.path.exists(src):
            return
        backup_dir = get_backup_dir()
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        dest = os.path.join(backup_dir, f'attendance_{timestamp}.db')
        shutil.copy2(src, dest)
        print(f"Backup successful: {dest}")
    except Exception as e:
        print(f"Backup failed: {e}")

def schedule_backup():
    now = datetime.datetime.now()
    target = now.replace(hour=12, minute=0, second=0, microsecond=0)
    if now >= target:
        target += datetime.timedelta(days=1)
    delta = (target - now).total_seconds()
    threading.Timer(delta, backup_and_reschedule).start()

def backup_and_reschedule():
    backup_db()
    schedule_backup()

# ---------- Database Setup ----------
def init_db():
    try:
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
                status TEXT,
                leave_reason TEXT,
                shift_code TEXT,
                FOREIGN KEY (staff_id) REFERENCES staff(staff_id),
                UNIQUE(staff_id, date)
            )
        ''')
        # Add columns if missing
        c.execute("PRAGMA table_info(attendance)")
        columns = [col[1] for col in c.fetchall()]
        if 'leave_reason' not in columns:
            c.execute("ALTER TABLE attendance ADD COLUMN leave_reason TEXT")
        if 'shift_code' not in columns:
            c.execute("ALTER TABLE attendance ADD COLUMN shift_code TEXT")
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
        c.execute('''
            CREATE TABLE IF NOT EXISTS config (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        c.execute("INSERT OR IGNORE INTO config (key, value) VALUES ('admin_password', 'admin123')")
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        messagebox.showerror("Database Error", f"Failed to initialize database:\n{str(e)}")
        return False

# ---------- Database Helpers ----------
def get_admin_password():
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT value FROM config WHERE key='admin_password'")
        row = c.fetchone()
        conn.close()
        return row[0] if row else 'admin123'
    except:
        return 'admin123'

def set_admin_password(new_password):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE config SET value=? WHERE key='admin_password'", (new_password,))
        conn.commit()
        conn.close()
        return True
    except:
        return False

def check_password(input_pw):
    return input_pw == get_admin_password()

def verify_password(parent=None):
    pw = simpledialog.askstring("Password Required", "Enter admin password:", show='*', parent=parent)
    if pw is None:
        return False
    return check_password(pw)

def get_staff(staff_id):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT staff_id, name, batch FROM staff WHERE staff_id=?", (staff_id,))
        row = c.fetchone()
        conn.close()
        return row
    except Exception as e:
        write_log_to_file(f"get_staff error: {e}")
        return None

def get_staff_by_name(name):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT staff_id FROM staff WHERE name=?", (name,))
        row = c.fetchone()
        conn.close()
        return row[0] if row else None
    except Exception as e:
        write_log_to_file(f"get_staff_by_name error: {e}")
        return None

def get_today_attendance(staff_id):
    try:
        today = datetime.date.today().isoformat()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT checkin, checkout, status, leave_reason, shift_code FROM attendance WHERE staff_id=? AND date=?", (staff_id, today))
        row = c.fetchone()
        conn.close()
        return row
    except Exception as e:
        write_log_to_file(f"get_today_attendance error: {e}")
        return None

def set_checkin(staff_id, time_str, shift_code=''):
    try:
        today = datetime.date.today().isoformat()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT status FROM attendance WHERE staff_id=? AND date=?", (staff_id, today))
        row = c.fetchone()
        existing_status = row[0] if row else ''
        tags = [t.strip() for t in existing_status.split(',') if t.strip()]
        if 'Clocked In' not in tags:
            tags.append('Clocked In')
        new_status = ', '.join(tags) if tags else 'Clocked In'
        c.execute('''
            INSERT OR REPLACE INTO attendance (staff_id, date, checkin, checkout, status, leave_reason, shift_code)
            VALUES (?, ?, ?, NULL, ?, ?, ?)
        ''', (staff_id, today, time_str, new_status, '', shift_code))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        write_log_to_file(f"set_checkin error: {e}")
        return False

def set_checkout(staff_id, time_str, leave_reason=''):
    try:
        today = datetime.date.today().isoformat()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT checkin, status FROM attendance WHERE staff_id=? AND date=?", (staff_id, today))
        row = c.fetchone()
        if not row:
            conn.close()
            return False
        checkin, status = row
        if not checkin:
            conn.close()
            return False

        schedule = get_work_schedule_for_date(staff_id, today)
        if schedule:
            work_start, work_end = schedule
        else:
            work_start, work_end = WORK_START, WORK_END
        work_start = pad_time(work_start)
        work_end = pad_time(work_end)

        tags = [t.strip() for t in status.split(',') if t.strip()] if status else []

        try:
            ci_dt = datetime.datetime.strptime(checkin, "%H:%M:%S")
            ws_dt = datetime.datetime.strptime(work_start, "%H:%M:%S")
            co_dt = datetime.datetime.strptime(time_str, "%H:%M:%S")
            we_dt = datetime.datetime.strptime(work_end, "%H:%M:%S")

            if we_dt < ws_dt:  # night shift
                if co_dt < ci_dt:
                    co_abs = co_dt + datetime.timedelta(days=1)
                else:
                    co_abs = co_dt
                we_abs = we_dt + datetime.timedelta(days=1)
                checkin_dev = (ci_dt - ws_dt).total_seconds() / 60.0
                checkout_dev = (co_abs - we_abs).total_seconds() / 60.0
            else:
                checkin_dev = (ci_dt - ws_dt).total_seconds() / 60.0
                checkout_dev = (co_dt - we_dt).total_seconds() / 60.0

            if checkin_dev > 0:
                tags.append('Late')
            elif checkin_dev < 0:
                tags.append('Early In')
            else:
                pass

            if checkout_dev > 0:
                tags.append('Overtime')
            elif checkout_dev < 0:
                tags.append('Early Leave')
            else:
                pass

            if not any(t in ['Late', 'Early In', 'Overtime', 'Early Leave'] for t in tags):
                tags.append('Normal')

            tags = [t for t in tags if t != 'Clocked In']
            unique_tags = []
            for t in tags:
                if t not in unique_tags:
                    unique_tags.append(t)
            new_status = ', '.join(unique_tags)

            c.execute('''
                UPDATE attendance SET checkout=?, status=?, leave_reason=?
                WHERE staff_id=? AND date=?
            ''', (time_str, new_status, leave_reason, staff_id, today))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            write_log_to_file(f"set_checkout deviation error: {e}")
            conn.close()
            return False
    except Exception as e:
        write_log_to_file(f"set_checkout error: {e}")
        return False

def override_checkin(staff_id, time_str, shift_code=''):
    try:
        today = datetime.date.today().isoformat()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('''
            UPDATE attendance SET checkin=?, checkout=NULL, status=?, leave_reason='', shift_code=?
            WHERE staff_id=? AND date=?
        ''', (time_str, 'Override', shift_code, staff_id, today))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        write_log_to_file(f"override_checkin error: {e}")
        return False

def upsert_attendance(staff_id, date_str, checkin, checkout):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('''
            INSERT OR REPLACE INTO attendance (staff_id, date, checkin, checkout, status, leave_reason, shift_code)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (staff_id, date_str, checkin, checkout, '', '', ''))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        write_log_to_file(f"upsert_attendance error: {e}")
        return False

def upsert_work_schedule(staff_id, start_date, end_date, work_start, work_end):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('''
            INSERT OR REPLACE INTO work_schedule (staff_id, start_date, end_date, work_start, work_end)
            VALUES (?, ?, ?, ?, ?)
        ''', (staff_id, start_date, end_date, work_start, work_end))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        write_log_to_file(f"upsert_work_schedule error: {e}")
        return False

def get_work_schedule_for_date(staff_id, date_str):
    try:
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
    except Exception as e:
        write_log_to_file(f"get_work_schedule_for_date error: {e}")
        return None

def get_monthly_attendance(year, month):
    try:
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year+1}-01-01"
        else:
            end_date = f"{year}-{month+1:02d}-01"
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('''
            SELECT s.staff_id, s.name, s.batch, a.date, a.checkin, a.checkout, a.status, a.leave_reason, a.shift_code
            FROM attendance a
            JOIN staff s ON a.staff_id = s.staff_id
            WHERE a.date >= ? AND a.date < ?
            ORDER BY s.staff_id, a.date
        ''', (start_date, end_date))
        rows = c.fetchall()
        conn.close()
        return rows
    except Exception as e:
        write_log_to_file(f"get_monthly_attendance error: {e}")
        return []

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
        self.root.title("Staff Attendance System v5.7")
        self.root.geometry("900x700")
        self.show_db_path()
        self.confirm_dialog = None
        self.current_staff_id = None
        self.current_name = None
        self.current_batch = None
        self.create_widgets()
        self.barcode_entry.bind("<Return>", self.on_barcode_scan)
        self.update_status()
        schedule_backup()

    def show_db_path(self):
        messagebox.showinfo("Database Location",
                            f"Attendance records stored at:\n{DB_PATH}\n\n"
                            f"Standard work hours: {STANDARD_HOURS} hrs (exact)\n"
                            "Missing clock-in/out will be marked in reports.\n"
                            "Auto backup at 12:00 daily.\n"
                            "Daily activity log stored in backup folder.")

    def create_widgets(self):
        top_frame = ttk.LabelFrame(self.root, text="Scan Barcode", padding=10)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(top_frame, text="Scan / Enter Staff ID:", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.barcode_entry = ttk.Entry(top_frame, width=40, font=("Arial", 14))
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
        # Removed Import Roster and Shift Code Reference buttons
        ttk.Button(btn_frame, text="Monthly Exceptions", command=self.show_monthly_summary).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Full Monthly Report", command=self.export_full_monthly_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="View Daily Log", command=self.view_daily_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Change Password", command=self.change_password).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Exit", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

        log_frame = ttk.LabelFrame(self.root, text="Recent Activity", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log_message(self, msg):
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        full_msg = f"{timestamp} - {msg}\n"
        self.log_text.insert(tk.END, full_msg)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        write_log_to_file(msg)

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
                    status = att[2]
                    leave_reason = att[3]
                    shift_code = att[4]
                    if checkout_time:
                        reason_display = f" | Leave: {leave_reason}" if leave_reason else ""
                        shift_display = f" | Shift: {shift_code}" if shift_code else ""
                        self.status_var.set(f"Clocked out at {checkout_time} | Status: {status}{reason_display}{shift_display}")
                    else:
                        shift_display = f" | Shift: {shift_code}" if shift_code else ""
                        self.status_var.set(f"Clocked in at {checkin_time} | Status: {status}{shift_display}")
                else:
                    self.status_var.set("Not clocked in today")
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

    def view_daily_log(self):
        log_path = get_today_log_path()
        if not os.path.exists(log_path):
            messagebox.showinfo("No Log", "No log entries for today yet.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Daily Log - {datetime.date.today().isoformat()}")
        win.geometry("700x500")

        frame = ttk.Frame(win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget = tk.Text(frame, wrap=tk.NONE, font=("Courier", 10))
        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=text_widget.xview)
        text_widget.config(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        try:
            with open(log_path, 'r', encoding='utf-8') as f:
                content = f.read()
            text_widget.insert(tk.END, content)
            text_widget.config(state=tk.DISABLED)
        except Exception as e:
            text_widget.insert(tk.END, f"Error reading log: {str(e)}")

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Open Log Folder",
                   command=lambda: os.startfile(os.path.dirname(log_path)) if os.name == 'nt' else \
                       messagebox.showinfo("Info", f"Log folder: {os.path.dirname(log_path)}")).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side=tk.LEFT, padx=5)

    # ---------- Password Management ----------
    def change_password(self):
        if not verify_password(self.root):
            messagebox.showerror("Error", "Incorrect password.")
            return
        new_pw = simpledialog.askstring("Change Password", "Enter new password:", show='*', parent=self.root)
        if new_pw:
            confirm = simpledialog.askstring("Change Password", "Confirm new password:", show='*', parent=self.root)
            if new_pw == confirm:
                if set_admin_password(new_pw):
                    messagebox.showinfo("Success", "Password updated successfully.")
                else:
                    messagebox.showerror("Error", "Failed to update password.")
            else:
                messagebox.showerror("Error", "Passwords do not match.")

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
                action = "Clock-in (override)"
                action_key = "override"
            else:
                action = "Clock-out"
                action_key = "checkout"
        else:
            action = "Clock-in"
            action_key = "checkin"

        self.show_confirmation(staff_id, name, batch, action, action_key, now)

    # ---------- Confirmation Dialog ----------
    def show_confirmation(self, staff_id, name, batch, action, action_key, current_time):
        dialog = tk.Toplevel(self.root)
        dialog.title("Confirm Attendance")
        dialog.geometry("650x550")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()

        self.barcode_entry.config(state=tk.DISABLED)
        self.scan_btn.config(state=tk.DISABLED)
        self.confirm_dialog = dialog

        # Countdown only for clock-out/override? No, override is like clock-in, so no countdown. Clock-out has countdown.
        if action_key == "checkout":
            countdown = 30
        else:
            countdown = None

        timer_id = None

        # Staff info
        ttk.Label(dialog, text="Staff:", font=("Arial", 12)).grid(row=0, column=0, padx=15, pady=8, sticky=tk.W)
        ttk.Label(dialog, text=f"{name} ({staff_id})", font=("Arial", 12, "bold")).grid(row=0, column=1, padx=15, pady=8, sticky=tk.W)

        ttk.Label(dialog, text="Batch:", font=("Arial", 12)).grid(row=1, column=0, padx=15, pady=8, sticky=tk.W)
        ttk.Label(dialog, text=batch or "-", font=("Arial", 12)).grid(row=1, column=1, padx=15, pady=8, sticky=tk.W)

        ttk.Label(dialog, text="Current time:", font=("Arial", 12)).grid(row=2, column=0, padx=15, pady=8, sticky=tk.W)
        ttk.Label(dialog, text=current_time, font=("Arial", 12, "bold")).grid(row=2, column=1, padx=15, pady=8, sticky=tk.W)

        ttk.Label(dialog, text="Action:", font=("Arial", 12)).grid(row=3, column=0, padx=15, pady=8, sticky=tk.W)
        ttk.Label(dialog, text=action, font=("Arial", 12, "bold"), foreground="blue").grid(row=3, column=1, padx=15, pady=8, sticky=tk.W)

        today = datetime.date.today().isoformat()

        # ---- Shift Selection (only for clock-in and override) ----
        shift_combo = None
        if action_key in ("checkin", "override"):
            # Get current schedule if exists
            existing_schedule = get_work_schedule_for_date(staff_id, today)
            shift_list = []
            for code, times in SHIFT_MAP.items():
                if times is not None:
                    start, end = times
                    shift_list.append(f"{code}: {start} - {end}")

            current_code = None
            if existing_schedule:
                start, end = existing_schedule
                for code, times in SHIFT_MAP.items():
                    if times and times[0] == start and times[1] == end:
                        current_code = code
                        break

            ttk.Label(dialog, text="Select Shift:", font=("Arial", 12)).grid(row=4, column=0, padx=15, pady=8, sticky=tk.W)
            if shift_list:
                combo = ttk.Combobox(dialog, values=shift_list, width=40, font=("Arial", 10))
                combo.grid(row=4, column=1, padx=15, pady=8)
                if current_code:
                    default_text = f"{current_code}: {existing_schedule[0]} - {existing_schedule[1]}"
                    if default_text in shift_list:
                        combo.set(default_text)
                    else:
                        combo.current(0)
                else:
                    combo.current(0)
                shift_combo = combo
            else:
                ttk.Label(dialog, text="No shifts defined (using default hours).", font=("Arial", 12), foreground="red").grid(row=4, column=1, padx=15, pady=8)
                self.log_message("Warning: No shifts in SHIFT_MAP, using default hours.")
        else:
            # For clock-out, display current schedule (read-only) and no combo
            existing_schedule = get_work_schedule_for_date(staff_id, today)
            if existing_schedule:
                start, end = existing_schedule
                code = None
                for c, times in SHIFT_MAP.items():
                    if times and times[0] == start and times[1] == end:
                        code = c
                        break
                if code:
                    display_text = f"{code}: {start} - {end}"
                else:
                    display_text = f"{start} - {end}"
                ttk.Label(dialog, text="Current Schedule:", font=("Arial", 12)).grid(row=4, column=0, padx=15, pady=8, sticky=tk.W)
                ttk.Label(dialog, text=display_text, font=("Arial", 12, "bold"), foreground="green").grid(row=4, column=1, padx=15, pady=8, sticky=tk.W)
            else:
                ttk.Label(dialog, text="No schedule set.", font=("Arial", 12), foreground="orange").grid(row=4, column=0, columnspan=2, padx=15, pady=8, sticky=tk.W)

        # ---- Leave Reason (only for clock-out/override? Override is like clock-in, but may also need leave reason? We'll keep for checkout only) ----
        leave_vars = []
        leave_frame = None
        if action_key == "checkout":
            leave_frame = ttk.LabelFrame(dialog, text="Early Leave Reasons", padding=10)
            # For checkout, we place it after the schedule row (row 5)
            leave_frame.grid(row=5, column=0, columnspan=2, padx=15, pady=8, sticky="ew")
            reasons = ["CO", "Annual Leave", "Sick Leave", "Other"]
            for i, reason in enumerate(reasons):
                var = tk.BooleanVar()
                cb = ttk.Checkbutton(leave_frame, text=reason, variable=var)
                cb.grid(row=i, column=0, padx=5, pady=2, sticky=tk.W)
                leave_vars.append((reason, var))

        # Countdown label
        countdown_label = None
        if countdown is not None:
            countdown_label = ttk.Label(dialog, text=f"Auto‑confirm in {countdown} seconds", font=("Arial", 11))
            countdown_label.grid(row=6, column=0, columnspan=2, pady=15)
        else:
            # For clock-in, no countdown, but we need to place the countdown label row empty or shift buttons up?
            # We'll leave a spacer or just place buttons at row 6.
            pass

        # Buttons row
        btn_row = 6 if countdown is not None else 5  # if no countdown, buttons go to row 5 (since leave_frame might be at row 5)
        # Actually adjust: if leave_frame exists, it's row 5, countdown row 6, buttons row 7. If no leave_frame and no countdown, buttons row 5.
        # Let's use dynamic positioning.
        row_counter = 4  # after staff info
        # Schedule row
        row_counter += 1
        # Leave reason row (if checkout)
        if leave_frame:
            row_counter += 1
        # Countdown row (if any)
        if countdown is not None:
            row_counter += 1
        # Buttons row
        btn_row = row_counter + 1

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=btn_row, column=0, columnspan=2, pady=20)

        def reset_after_dialog():
            self.barcode_entry.config(state=tk.NORMAL)
            self.scan_btn.config(state=tk.NORMAL)
            self.barcode_entry.focus_set()
            self.confirm_dialog = None
            if timer_id:
                dialog.after_cancel(timer_id)

        def get_leave_reason():
            if leave_vars:
                selected = [reason for reason, var in leave_vars if var.get()]
                return ', '.join(selected) if selected else ''
            return ''

        def perform_action(selected_shift=None, leave_reason=''):
            # Update schedule if shift selected (only for clock-in/override)
            if selected_shift and selected_shift in SHIFT_MAP and SHIFT_MAP[selected_shift] is not None:
                start, end = SHIFT_MAP[selected_shift]
                if upsert_work_schedule(staff_id, today, today, start, end):
                    self.log_message(f"Updated shift to {selected_shift} ({start}-{end}) for today")
                else:
                    self.log_message("Failed to update shift schedule")

            # Main action
            success = False
            try:
                if action_key == "checkin":
                    success = set_checkin(staff_id, current_time, selected_shift or '')
                    if success:
                        self.log_message(f"Clocked in at {current_time} | Shift: {selected_shift or 'None'}")
                    else:
                        self.log_message("Failed to record clock-in (database error)")
                elif action_key == "checkout":
                    success = set_checkout(staff_id, current_time, leave_reason)
                    if success:
                        self.log_message(f"Clocked out at {current_time} | Leave Reason: {leave_reason or 'None'}")
                    else:
                        self.log_message("Failed to record clock-out (database error)")
                elif action_key == "override":
                    success = override_checkin(staff_id, current_time, selected_shift or '')
                    if success:
                        self.log_message(f"Overrode clock-in at {current_time} (previous checkout cleared) | Shift: {selected_shift or 'None'}")
                    else:
                        self.log_message("Failed to override clock-in (database error)")
                else:
                    self.log_message("Unknown action – nothing stored")
                    success = False
            except Exception as e:
                self.log_message(f"Error during action: {str(e)}")
                success = False

            if success:
                self.update_status(staff_id)

        def do_confirm():
            nonlocal timer_id
            if timer_id:
                dialog.after_cancel(timer_id)
                timer_id = None
            selected_shift = None
            if shift_combo:
                selected_text = shift_combo.get()
                if selected_text:
                    code = selected_text.split(':')[0].strip()
                    selected_shift = code
            leave_reason = get_leave_reason() if leave_frame else ''
            perform_action(selected_shift, leave_reason)
            reset_after_dialog()
            dialog.destroy()

        def do_cancel():
            nonlocal timer_id
            if timer_id:
                dialog.after_cancel(timer_id)
                timer_id = None
            self.log_message(f"Confirmation cancelled for {name}")
            reset_after_dialog()
            dialog.destroy()

        confirm_btn = ttk.Button(btn_frame, text="Confirm", command=do_confirm, width=12)
        confirm_btn.pack(side=tk.LEFT, padx=15)
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=do_cancel, width=12)
        cancel_btn.pack(side=tk.LEFT, padx=15)

        def update_countdown():
            nonlocal countdown, timer_id
            if countdown is None:
                return
            if countdown <= 0:
                if timer_id:
                    timer_id = None
                selected_shift = None
                if shift_combo:
                    selected_text = shift_combo.get()
                    if selected_text:
                        code = selected_text.split(':')[0].strip()
                        selected_shift = code
                leave_reason = get_leave_reason() if leave_frame else ''
                perform_action(selected_shift, leave_reason)
                reset_after_dialog()
                dialog.destroy()
                return

            countdown_label.config(text=f"Auto‑confirm in {countdown} seconds")
            countdown -= 1
            timer_id = dialog.after(1000, update_countdown)

        if countdown is not None:
            timer_id = dialog.after(1000, update_countdown)

        dialog.protocol("WM_DELETE_WINDOW", do_cancel)

    # ---------- Staff Management ----------
    def manage_staff(self):
        if not verify_password(self.root):
            messagebox.showerror("Error", "Incorrect password.")
            return
        self._open_manage_staff()

    def _open_manage_staff(self):
        win = tk.Toplevel(self.root)
        win.title("Manage Staff")
        win.geometry("550x450")

        tree = ttk.Treeview(win, columns=("ID", "Name", "Batch"), show="headings")
        tree.heading("ID", text="Staff ID")
        tree.heading("Name", text="Name")
        tree.heading("Batch", text="Batch")
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        def refresh_staff_list():
            for row in tree.get_children():
                tree.delete(row)
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
                for r in c.fetchall():
                    tree.insert("", tk.END, values=r)
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load staff: {str(e)}")

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
                    try:
                        conn = sqlite3.connect(DB_PATH)
                        c = conn.cursor()
                        c.execute("INSERT INTO staff (staff_id, name, batch) VALUES (?, ?, ?)",
                                  (staff_id, name, batch))
                        conn.commit()
                        conn.close()
                        refresh_staff_list()
                        self.log_message(f"Added staff: {name} ({staff_id})")
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to add staff: {str(e)}")

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
                try:
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    c.execute("UPDATE staff SET name=?, batch=? WHERE staff_id=?", (new_name, new_batch, staff_id))
                    conn.commit()
                    conn.close()
                    refresh_staff_list()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to update staff: {str(e)}")

        def delete_staff():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("Info", "Select a staff to delete.")
                return
            values = tree.item(selected[0])['values']
            staff_id = values[0]
            if messagebox.askyesno("Delete", f"Delete staff '{values[1]}' and all attendance records?"):
                try:
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    c.execute("DELETE FROM attendance WHERE staff_id=?", (staff_id,))
                    c.execute("DELETE FROM staff WHERE staff_id=?", (staff_id,))
                    conn.commit()
                    conn.close()
                    refresh_staff_list()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to delete staff: {str(e)}")

        def export_staff():
            try:
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv")],
                    title="Export Staff List"
                )
                if not file_path:
                    return
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
                rows = c.fetchall()
                conn.close()
                if not rows:
                    messagebox.showinfo("No Data", "No staff records to export.")
                    return
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Staff ID", "Name", "Batch"])
                    writer.writerows(rows)
                messagebox.showinfo("Export Complete", f"Staff list exported to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")

        ttk.Button(btn_frame, text="Add", command=add_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Edit", command=edit_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", command=delete_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Export Staff", command=export_staff).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side=tk.RIGHT, padx=5)

    # ---------- Add New Staff ----------
    def add_new_staff(self, staff_id):
        name = simpledialog.askstring("Add Staff", "Enter staff name:", parent=self.root)
        if name:
            batch = simpledialog.askstring("Add Staff", "Enter batch (optional):", parent=self.root)
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("INSERT INTO staff (staff_id, name, batch) VALUES (?, ?, ?)",
                          (staff_id, name, batch))
                conn.commit()
                conn.close()
                self.log_message(f"Added staff: {name} ({staff_id})")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add staff: {str(e)}")

    # ---------- Monthly Exception Report ----------
    def show_monthly_summary(self):
        if not verify_password(self.root):
            messagebox.showerror("Error", "Incorrect password.")
            return
        self._open_monthly_summary()

    def _open_monthly_summary(self):
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

            all_staff = []
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
                all_staff = c.fetchall()
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load staff: {str(e)}")
                return

            if not all_staff:
                messagebox.showinfo("No Staff", "No staff records found.")
                return

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

            exceptions = {}

            for staff_id, name, batch in all_staff:
                try:
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    c.execute('''
                        SELECT date, checkin, checkout, status, leave_reason, shift_code
                        FROM attendance
                        WHERE staff_id=? AND date >= ? AND date < ?
                    ''', (staff_id, first_day.isoformat(), next_month.isoformat()))
                    records = {row[0]: (row[1], row[2], row[3], row[4], row[5]) for row in c.fetchall()}
                    conn.close()
                except Exception as e:
                    self.log_message(f"Error loading attendance for {name}: {e}")
                    continue

                for date_str in all_dates:
                    schedule = get_work_schedule_for_date(staff_id, date_str)
                    if schedule:
                        work_start, work_end = schedule
                    else:
                        work_start, work_end = WORK_START, WORK_END

                    if date_str in records:
                        checkin, checkout, status, leave_reason, shift_code = records[date_str]
                        if status:
                            status_lower = status.lower()
                            if 'late' in status_lower or 'early leave' in status_lower or 'forgot' in status_lower or 'no checkout' in status_lower or 'no checkin' in status_lower:
                                exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                                exceptions[staff_id]['days'].append((date_str, status, f"Checkin: {checkin}, Checkout: {checkout}", leave_reason, shift_code))
                    else:
                        if schedule:
                            exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                            exceptions[staff_id]['days'].append((date_str, "Forgot Check", "No check-in/out recorded", "", ""))

            if not exceptions:
                messagebox.showinfo("All Good", "No exceptions found for this month.")
                return

            result_win = tk.Toplevel(win)
            result_win.title(f"Exception Report - {year}-{month:02d}")
            result_win.geometry("1000x500")

            tree = ttk.Treeview(result_win, columns=("Staff", "Batch", "Date", "Issue", "Detail", "Leave Reason", "Shift Code"), show="headings")
            tree.heading("Staff", text="Staff (ID)")
            tree.heading("Batch", text="Batch")
            tree.heading("Date", text="Date")
            tree.heading("Issue", text="Issue Type")
            tree.heading("Detail", text="Detail")
            tree.heading("Leave Reason", text="Leave Reason")
            tree.heading("Shift Code", text="Shift Code")
            tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            for staff_id, data in exceptions.items():
                for date_str, issue, detail, leave_reason, shift_code in data['days']:
                    tree.insert("", tk.END, values=(
                        f"{data['name']} ({staff_id})",
                        data['batch'] or "-",
                        date_str,
                        issue,
                        detail,
                        leave_reason or "-",
                        shift_code or "-"
                    ))

            def export_exceptions():
                file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                         filetypes=[("CSV files", "*.csv")])
                if not file_path:
                    return
                try:
                    with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                        writer = csv.writer(f)
                        writer.writerow(["Staff ID", "Name", "Batch", "Date", "Issue Type", "Detail", "Leave Reason", "Shift Code"])
                        for staff_id, data in exceptions.items():
                            for date_str, issue, detail, leave_reason, shift_code in data['days']:
                                writer.writerow([staff_id, data['name'], data['batch'] or "", date_str, issue, detail, leave_reason, shift_code])
                    messagebox.showinfo("Export", f"Report exported to {file_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"Export failed: {str(e)}")

            ttk.Button(result_win, text="Export CSV", command=export_exceptions).pack(pady=5)

        ttk.Button(win, text="Generate", command=generate).pack(pady=20)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=5)

    # ---------- Full Monthly Report ----------
    def export_full_monthly_report(self):
        if not verify_password(self.root):
            messagebox.showerror("Error", "Incorrect password.")
            return
        self._open_full_monthly_report()

    def _open_full_monthly_report(self):
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

            all_staff = []
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
                all_staff = c.fetchall()
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load staff: {str(e)}")
                return

            if not all_staff:
                messagebox.showinfo("No Staff", "No staff records found.")
                return

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
                try:
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    c.execute('''
                        SELECT date, checkin, checkout, status, leave_reason, shift_code
                        FROM attendance
                        WHERE staff_id=? AND date >= ? AND date < ?
                    ''', (staff_id, first_day.isoformat(), next_month.isoformat()))
                    records = {row[0]: (row[1], row[2], row[3], row[4], row[5]) for row in c.fetchall()}
                    conn.close()
                except Exception as e:
                    self.log_message(f"Error loading attendance for {name}: {e}")
                    continue

                for date_str in all_dates:
                    schedule = get_work_schedule_for_date(staff_id, date_str)
                    if schedule:
                        work_start, work_end = schedule
                    else:
                        work_start, work_end = WORK_START, WORK_END

                    work_start = pad_time(work_start)
                    work_end = pad_time(work_end)

                    if date_str in records:
                        checkin, checkout, status, leave_reason, shift_code = records[date_str]
                        if checkin and checkout:
                            work_hrs = calculate_work_hours(checkin, checkout)
                            try:
                                ci_dt = datetime.datetime.strptime(checkin, "%H:%M:%S")
                                co_dt = datetime.datetime.strptime(checkout, "%H:%M:%S")
                                ws_dt = datetime.datetime.strptime(work_start, "%H:%M:%S")
                                we_dt = datetime.datetime.strptime(work_end, "%H:%M:%S")
                                if we_dt < ws_dt:
                                    if co_dt < ci_dt:
                                        co_abs = co_dt + datetime.timedelta(days=1)
                                    else:
                                        co_abs = co_dt
                                    we_abs = we_dt + datetime.timedelta(days=1)
                                    checkin_dev = (ci_dt - ws_dt).total_seconds() / 60.0
                                    checkout_dev = (co_abs - we_abs).total_seconds() / 60.0
                                else:
                                    checkin_dev = (ci_dt - ws_dt).total_seconds() / 60.0
                                    checkout_dev = (co_dt - we_dt).total_seconds() / 60.0
                                checkin_dev = round(checkin_dev)
                                checkout_dev = round(checkout_dev)
                            except:
                                checkin_dev = None
                                checkout_dev = None

                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Checkin": checkin,
                                "Checkout": checkout,
                                "Work Hours": f"{work_hrs:.2f}",
                                "Status": status or "Normal",
                                "Checkin Deviation (min)": checkin_dev if checkin_dev is not None else "-",
                                "Checkout Deviation (min)": checkout_dev if checkout_dev is not None else "-",
                                "Leave Reason": leave_reason or "",
                                "Shift Code": shift_code or "",
                            })
                        elif checkin and not checkout:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Checkin": checkin,
                                "Checkout": "",
                                "Work Hours": "0.00",
                                "Status": status or "No Checkout Record",
                                "Checkin Deviation (min)": "-",
                                "Checkout Deviation (min)": "-",
                                "Leave Reason": "",
                                "Shift Code": shift_code or "",
                            })
                        elif not checkin and checkout:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Checkin": "",
                                "Checkout": checkout,
                                "Work Hours": "0.00",
                                "Status": status or "No Checkin Record",
                                "Checkin Deviation (min)": "-",
                                "Checkout Deviation (min)": "-",
                                "Leave Reason": "",
                                "Shift Code": shift_code or "",
                            })
                        else:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Checkin": "",
                                "Checkout": "",
                                "Work Hours": "0.00",
                                "Status": status or "Forgot Check",
                                "Checkin Deviation (min)": "-",
                                "Checkout Deviation (min)": "-",
                                "Leave Reason": "",
                                "Shift Code": "",
                            })
                    else:
                        if schedule:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Checkin": "",
                                "Checkout": "",
                                "Work Hours": "0.00",
                                "Status": "Forgot Check",
                                "Checkin Deviation (min)": "-",
                                "Checkout Deviation (min)": "-",
                                "Leave Reason": "",
                                "Shift Code": "",
                            })
                        else:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Checkin": "",
                                "Checkout": "",
                                "Work Hours": "0.00",
                                "Status": "Off Day",
                                "Checkin Deviation (min)": "-",
                                "Checkout Deviation (min)": "-",
                                "Leave Reason": "",
                                "Shift Code": "",
                            })

            if not report_data:
                messagebox.showinfo("No Data", "No data to export.")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Save Full Monthly Report"
            )
            if not file_path:
                return

            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    fieldnames = ["Staff ID", "Name", "Batch", "Date", "Checkin", "Checkout",
                                  "Work Hours", "Status", "Checkin Deviation (min)", "Checkout Deviation (min)",
                                  "Leave Reason", "Shift Code"]
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(report_data)
                messagebox.showinfo("Export Complete", f"Full monthly report saved to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")

        ttk.Button(win, text="Generate & Export CSV", command=generate).pack(pady=20)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=5)


# ---------- Main ----------
def main():
    try:
        if not init_db():
            return
        root = tk.Tk()
        app = AttendanceApp(root)
        root.mainloop()
    except Exception as e:
        error_msg = f"Unhandled exception:\n{str(e)}\n\n{traceback.format_exc()}"
        try:
            messagebox.showerror("Fatal Error", error_msg)
        except:
            print(error_msg)
        sys.exit(1)

if __name__ == "__main__":
    main()
