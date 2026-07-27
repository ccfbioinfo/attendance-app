"""
Staff Duty Attendance System v5.0
- Barcode scan with shift selection
- Clock-on / Clock-off (instead of check-in/out)
- 8.8 hrs threshold with 5-min grace
- Fixed shift options (selectable per day)
"""

import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import os
import sys
import csv

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ---------- Configuration ----------
STANDARD_HOURS = 8.8
GRACE_MINUTES = 5
GRACE_HOURS = GRACE_MINUTES / 60.0

# ---------- Shift Options (display, start, end) ----------
SHIFT_OPTIONS = [
    ("08:00 - 16:48", "08:00", "16:48"),
    ("08:30 - 17:18", "08:30", "17:18"),
    ("09:00 - 17:00", "09:00", "17:00"),
    ("09:00 - 17:48", "09:00", "17:48"),
    ("09:00 - 18:00", "09:00", "18:00"),
    ("10:00 - 18:48", "10:00", "18:48"),
    ("13:00 - 21:48", "13:00", "21:48"),
    ("21:30 - 08:30", "21:30", "08:30"),
]

# ---------- Database Path: Always next to .exe ----------
def get_db_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, 'attendance.db')

DB_PATH = get_db_path()

# ---------- Database Setup with Shift Columns ----------
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
            clock_on TEXT,
            clock_off TEXT,
            shift_start TEXT,
            shift_end TEXT,
            FOREIGN KEY (staff_id) REFERENCES staff(staff_id),
            UNIQUE(staff_id, date)
        )
    ''')
    # Check if shift columns exist (for upgrading old DB)
    c.execute("PRAGMA table_info(attendance)")
    columns = [col[1] for col in c.fetchall()]
    if 'shift_start' not in columns:
        c.execute("ALTER TABLE attendance ADD COLUMN shift_start TEXT")
    if 'shift_end' not in columns:
        c.execute("ALTER TABLE attendance ADD COLUMN shift_end TEXT")
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
    c.execute("SELECT clock_on, clock_off, shift_start, shift_end FROM attendance WHERE staff_id=? AND date=?", (staff_id, today))
    row = c.fetchone()
    conn.close()
    return row

def set_clock_on(staff_id, time_str, shift_start, shift_end):
    today = datetime.date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        INSERT OR REPLACE INTO attendance (staff_id, date, clock_on, clock_off, shift_start, shift_end)
        VALUES (?, ?, ?, NULL, ?, ?)
    ''', (staff_id, today, time_str, shift_start, shift_end))
    conn.commit()
    conn.close()

def set_clock_off(staff_id, time_str):
    today = datetime.date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        UPDATE attendance SET clock_off=?
        WHERE staff_id=? AND date=? AND clock_on IS NOT NULL
    ''', (time_str, staff_id, today))
    conn.commit()
    conn.close()

def upsert_attendance(staff_id, date_str, clock_on, clock_off, shift_start=None, shift_end=None):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    if shift_start and shift_end:
        c.execute('''
            INSERT OR REPLACE INTO attendance (staff_id, date, clock_on, clock_off, shift_start, shift_end)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (staff_id, date_str, clock_on, clock_off, shift_start, shift_end))
    else:
        c.execute('''
            INSERT OR REPLACE INTO attendance (staff_id, date, clock_on, clock_off, shift_start, shift_end)
            VALUES (?, ?, ?, ?, NULL, NULL)
        ''', (staff_id, date_str, clock_on, clock_off))
    conn.commit()
    conn.close()

def calculate_work_hours(clock_on_str, clock_off_str):
    try:
        ci = datetime.datetime.strptime(clock_on_str, "%H:%M:%S")
        co = datetime.datetime.strptime(clock_off_str, "%H:%M:%S")
        if co <= ci:
            co += datetime.timedelta(days=1)
        delta = co - ci
        return delta.total_seconds() / 3600.0
    except:
        return 0.0

def get_monthly_attendance(year, month):
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year+1}-01-01"
    else:
        end_date = f"{year}-{month+1:02d}-01"
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT s.staff_id, s.name, s.batch, a.date, a.clock_on, a.clock_off, a.shift_start, a.shift_end
        FROM attendance a
        JOIN staff s ON a.staff_id = s.staff_id
        WHERE a.date >= ? AND a.date < ?
          AND a.clock_on IS NOT NULL AND a.clock_off IS NOT NULL
        ORDER BY s.staff_id, a.date
    ''', (start_date, end_date))
    rows = c.fetchall()
    conn.close()
    return rows

# ---------- GUI Application ----------
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Staff Attendance System v5.0")
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
                            f"Standard daily hours: {STANDARD_HOURS} hrs (±{GRACE_MINUTES} min grace)\n"
                            "Select your shift after each scan.")

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
                    clock_on = att[0]
                    clock_off = att[1]
                    shift = f"{att[2]} - {att[3]}" if att[2] and att[3] else ""
                    if clock_off:
                        self.status_var.set(f"Clocked off at {clock_off} (Shift: {shift})")
                    else:
                        self.status_var.set(f"Clocked on at {clock_on} (Shift: {shift})")
                else:
                    self.status_var.set("Not clocked on today")
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
                    staff_id, date_str, clock_on, clock_off = row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip()
                    try:
                        datetime.datetime.strptime(date_str, "%Y-%m-%d")
                    except ValueError:
                        self.log_message(f"Skipping invalid date: {date_str}")
                        continue
                    if not get_staff(staff_id):
                        self.log_message(f"Staff {staff_id} not found, skipping.")
                        continue
                    upsert_attendance(staff_id, date_str, clock_on, clock_off)
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
                    clock_on = str(row[2]).strip() if row[2] else None
                    clock_off = str(row[3]).strip() if row[3] else None
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
                    upsert_attendance(staff_id, date_str, clock_on, clock_off)
                    count += 1
            self.log_message(f"Imported {count} attendance records from Excel.")
            messagebox.showinfo("Success", f"Imported {count} records.")
        except Exception as e:
            messagebox.showerror("Error", f"Excel import failed: {str(e)}")

    # ---------- Barcode Scan with Shift Selection ----------
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

        # Check today's attendance
        today_record = get_today_attendance(staff_id)
        now = datetime.datetime.now().strftime("%H:%M:%S")

        if today_record:
            clock_on = today_record[0]
            clock_off = today_record[1]
            if clock_off:
                # Already clocked off today
                messagebox.showinfo("Already Clocked Off", f"{name} already clocked off at {clock_off} today.\nNo action taken.")
                self.log_message(f"{name} already clocked off at {clock_off}")
                return
            else:
                # Clock off process: use existing shift, no need to select again
                shift_start = today_record[2]
                shift_end = today_record[3]
                # Confirm clock off
                if messagebox.askyesno("Clock Off", f"Clock off for {name} at {now}?\nShift: {shift_start} - {shift_end}"):
                    set_clock_off(staff_id, now)
                    work_hours = calculate_work_hours(clock_on, now)
                    self.log_message(f"{name} clocked off at {now}. Worked {work_hours:.2f} hrs.")
                    self.update_status(staff_id)
                else:
                    self.log_message("Clock off cancelled.")
                return
        else:
            # No record today -> Clock on with shift selection
            self.show_shift_selection(staff_id, name, batch, now)

    def show_shift_selection(self, staff_id, name, batch, current_time):
        """Popup for shift selection, then clock on."""
        win = tk.Toplevel(self.root)
        win.title("Select Shift")
        win.geometry("400x300")
        win.transient(self.root)
        win.grab_set()
        win.focus_force()

        ttk.Label(win, text=f"Staff: {name} ({staff_id})", font=("Arial", 12)).pack(pady=10)
        ttk.Label(win, text="Select your working shift for today:").pack(pady=5)

        # Determine recommended shift based on current time
        now_time = datetime.datetime.strptime(current_time, "%H:%M:%S").time()
        recommended_idx = 0
        for idx, (label, start, end) in enumerate(SHIFT_OPTIONS):
            start_time = datetime.datetime.strptime(start, "%H:%M").time()
            end_time = datetime.datetime.strptime(end, "%H:%M").time()
            # For night shift (start > end), we need to treat differently
            if start_time > end_time:
                # Night shift: if current time is between start and midnight, or between midnight and end
                if now_time >= start_time or now_time <= end_time:
                    recommended_idx = idx
                    break
            else:
                # Normal shift: if current time is within 30 minutes of start
                diff = abs((datetime.datetime.combine(datetime.date.today(), now_time) -
                            datetime.datetime.combine(datetime.date.today(), start_time)).total_seconds() / 60)
                if diff <= 30:
                    recommended_idx = idx
                    break
        # If no match, default to first

        # ComboBox for shift selection
        shift_display = [opt[0] for opt in SHIFT_OPTIONS]
        selected_var = tk.StringVar(value=shift_display[recommended_idx])
        combo = ttk.Combobox(win, textvariable=selected_var, values=shift_display, state="readonly", width=25)
        combo.pack(pady=10)

        def confirm_shift():
            selected_label = selected_var.get()
            # Find selected shift
            shift_start = None
            shift_end = None
            for label, start, end in SHIFT_OPTIONS:
                if label == selected_label:
                    shift_start = start
                    shift_end = end
                    break
            if not shift_start or not shift_end:
                messagebox.showerror("Error", "Invalid shift selection.")
                return
            # Record clock on
            set_clock_on(staff_id, current_time, shift_start, shift_end)
            self.log_message(f"{name} clocked on at {current_time} (Shift: {shift_start} - {shift_end})")
            self.update_status(staff_id)
            win.destroy()

        def cancel():
            self.log_message("Shift selection cancelled.")
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="Confirm", command=confirm_shift).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=cancel).pack(side=tk.LEFT, padx=10)

    # ---------- Staff Management ----------
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

    # ---------- Monthly Exception Report (using clock_on/off and shift) ----------
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

            # Get all staff
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
            all_staff = c.fetchall()
            conn.close()

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
                # Get attendance records for the month
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute('''
                    SELECT date, clock_on, clock_off, shift_start, shift_end
                    FROM attendance
                    WHERE staff_id=? AND date >= ? AND date < ?
                ''', (staff_id, first_day.isoformat(), next_month.isoformat()))
                records = {row[0]: (row[1], row[2], row[3], row[4]) for row in c.fetchall()}
                conn.close()

                for date_str in all_dates:
                    if date_str in records:
                        clock_on, clock_off, shift_start, shift_end = records[date_str]
                        if clock_on and not clock_off:
                            exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                            exceptions[staff_id]['days'].append((date_str, "No Clock-off Record", f"Clocked-on: {clock_on}"))
                        elif not clock_on and clock_off:
                            exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                            exceptions[staff_id]['days'].append((date_str, "No Clock-on Record", f"Clocked-off: {clock_off}"))
                        else:
                            # Both present -> check late/early leave based on shift
                            if shift_start and shift_end:
                                try:
                                    ci = datetime.datetime.strptime(clock_on, "%H:%M:%S").time()
                                    co = datetime.datetime.strptime(clock_off, "%H:%M:%S").time()
                                    ws = datetime.datetime.strptime(shift_start, "%H:%M").time()
                                    we = datetime.datetime.strptime(shift_end, "%H:%M").time()
                                    if ci > ws:
                                        exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                                        exceptions[staff_id]['days'].append((date_str, "Late", f"Clocked-on: {clock_on} (Shift start: {shift_start})"))
                                    if co < we:
                                        exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                                        exceptions[staff_id]['days'].append((date_str, "Early Leave", f"Clocked-off: {clock_off} (Shift end: {shift_end})"))
                                except:
                                    pass
                    else:
                        # No record at all -> Forgot Clock
                        exceptions.setdefault(staff_id, {'name': name, 'batch': batch, 'days': []})
                        exceptions[staff_id]['days'].append((date_str, "Forgot Clock", "No clock-on/off recorded"))

            if not exceptions:
                messagebox.showinfo("All Good", "No exceptions found for this month.")
                return

            result_win = tk.Toplevel(win)
            result_win.title(f"Exception Report - {year}-{month:02d}")
            result_win.geometry("900x500")

            tree = ttk.Treeview(result_win, columns=("Staff", "Batch", "Date", "Issue", "Detail"), show="headings")
            tree.heading("Staff", text="Staff (ID)")
            tree.heading("Batch", text="Batch")
            tree.heading("Date", text="Date")
            tree.heading("Issue", text="Issue Type")
            tree.heading("Detail", text="Detail")
            tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            for staff_id, data in exceptions.items():
                for date_str, issue, detail in data['days']:
                    tree.insert("", tk.END, values=(
                        f"{data['name']} ({staff_id})",
                        data['batch'] or "-",
                        date_str,
                        issue,
                        detail
                    ))

            def export_exceptions():
                file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                         filetypes=[("CSV files", "*.csv")])
                if not file_path:
                    return
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Staff ID", "Name", "Batch", "Date", "Issue Type", "Detail"])
                    for staff_id, data in exceptions.items():
                        for date_str, issue, detail in data['days']:
                            writer.writerow([staff_id, data['name'], data['batch'] or "", date_str, issue, detail])
                messagebox.showinfo("Export", f"Report exported to {file_path}")

            ttk.Button(result_win, text="Export CSV", command=export_exceptions).pack(pady=5)

        ttk.Button(win, text="Generate", command=generate).pack(pady=20)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=5)

    # ---------- Full Monthly Report (with shift and 8.8 threshold) ----------
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

            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("SELECT staff_id, name, batch FROM staff ORDER BY name")
            all_staff = c.fetchall()
            conn.close()

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
                # Get attendance records
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute('''
                    SELECT date, clock_on, clock_off, shift_start, shift_end
                    FROM attendance
                    WHERE staff_id=? AND date >= ? AND date < ?
                ''', (staff_id, first_day.isoformat(), next_month.isoformat()))
                records = {row[0]: (row[1], row[2], row[3], row[4]) for row in c.fetchall()}
                conn.close()

                for date_str in all_dates:
                    if date_str in records:
                        clock_on, clock_off, shift_start, shift_end = records[date_str]
                        if clock_on and clock_off:
                            work_hrs = calculate_work_hours(clock_on, clock_off)
                            if abs(work_hrs - STANDARD_HOURS) <= GRACE_HOURS:
                                status = "Normal"
                            elif work_hrs < STANDARD_HOURS - GRACE_HOURS:
                                status = "Early Leave"
                            else:
                                status = "Overtime"
                            shift_info = f"{shift_start} - {shift_end}" if shift_start and shift_end else "N/A"
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Clock On": clock_on,
                                "Clock Off": clock_off,
                                "Shift": shift_info,
                                "Work Hours": f"{work_hrs:.2f}",
                                "Status": status
                            })
                        elif clock_on and not clock_off:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Clock On": clock_on,
                                "Clock Off": "",
                                "Shift": f"{shift_start} - {shift_end}" if shift_start and shift_end else "N/A",
                                "Work Hours": "0.00",
                                "Status": "No Clock-off Record"
                            })
                        elif not clock_on and clock_off:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Clock On": "",
                                "Clock Off": clock_off,
                                "Shift": f"{shift_start} - {shift_end}" if shift_start and shift_end else "N/A",
                                "Work Hours": "0.00",
                                "Status": "No Clock-on Record"
                            })
                        else:
                            report_data.append({
                                "Staff ID": staff_id,
                                "Name": name,
                                "Batch": batch or "",
                                "Date": date_str,
                                "Clock On": "",
                                "Clock Off": "",
                                "Shift": "",
                                "Work Hours": "0.00",
                                "Status": "Forgot Clock"
                            })
                    else:
                        report_data.append({
                            "Staff ID": staff_id,
                            "Name": name,
                            "Batch": batch or "",
                            "Date": date_str,
                            "Clock On": "",
                            "Clock Off": "",
                            "Shift": "",
                            "Work Hours": "0.00",
                            "Status": "Forgot Clock"
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

            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=["Staff ID", "Name", "Batch", "Date", "Clock On", "Clock Off", "Shift", "Work Hours", "Status"])
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
