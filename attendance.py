"""
Staff Duty Attendance System - with 10‑second confirmation timer
Database path auto-detects writable location.
"""

import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import os
import sys

# ---------- Determine a writable database path ----------
def get_db_path():
    """Return a writable path for attendance.db."""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        base_dir = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # First try to use the executable's directory
    db_path = os.path.join(base_dir, "attendance.db")
    # Check if we can write there
    try:
        test_file = os.path.join(base_dir, "write_test.tmp")
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        # If we reach here, directory is writable
        return db_path
    except (OSError, PermissionError):
        # Fallback to user's AppData directory (Windows) or ~/.local (Linux/macOS)
        if os.name == 'nt':  # Windows
            appdata = os.getenv('APPDATA')
            if not appdata:
                appdata = os.path.expanduser('~/AppData/Roaming')
            db_dir = os.path.join(appdata, 'AttendanceSystem')
        else:  # macOS/Linux
            db_dir = os.path.expanduser('~/.local/share/AttendanceSystem')
        if not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
        return os.path.join(db_dir, "attendance.db")

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
    conn.commit()
    conn.close()

# ---------- Database Helpers (all use DB_PATH) ----------
def get_staff(staff_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT staff_id, name, batch FROM staff WHERE staff_id=?", (staff_id,))
    row = c.fetchone()
    conn.close()
    return row

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

def get_monthly_summary(year, month):
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year+1}-01-01"
    else:
        end_date = f"{year}-{month+1:02d}-01"
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT s.staff_id, s.name, s.batch,
               SUM(strftime('%s', a.checkout) - strftime('%s', a.checkin)) AS total_sec,
               COUNT(a.id) AS days_worked
        FROM attendance a
        JOIN staff s ON a.staff_id = s.staff_id
        WHERE a.date >= ? AND a.date < ?
          AND a.checkin IS NOT NULL AND a.checkout IS NOT NULL
        GROUP BY s.staff_id
        ORDER BY s.name
    ''', (start_date, end_date))
    rows = c.fetchall()
    conn.close()
    return rows

def get_daily_details(staff_id, year, month):
    start_date = f"{year}-{month:02d}-01"
    if month == 12:
        end_date = f"{year+1}-01-01"
    else:
        end_date = f"{year}-{month+1:02d}-01"
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT date, checkin, checkout,
               (strftime('%s', checkout) - strftime('%s', checkin)) / 3600.0 AS hours
        FROM attendance
        WHERE staff_id=? AND date >= ? AND date < ?
          AND checkin IS NOT NULL AND checkout IS NOT NULL
        ORDER BY date
    ''', (staff_id, start_date, end_date))
    rows = c.fetchall()
    conn.close()
    return rows

# ---------- GUI Application (unchanged except DB_PATH) ----------
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Staff Attendance System")
        self.root.geometry("600x500")

        self.confirm_dialog = None
        self.timer_id = None
        self.countdown = 10

        self.current_staff_id = None
        self.current_name = None
        self.current_batch = None

        self.create_widgets()
        self.barcode_entry.bind("<Return>", self.on_barcode_scan)
        self.update_status()

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
        ttk.Button(btn_frame, text="Monthly Summary", command=self.show_monthly_summary).pack(side=tk.LEFT, padx=5)
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

        def do_confirm():
            if self.timer_id:
                dialog.after_cancel(self.timer_id)
                self.timer_id = None
            self.perform_action(staff_id, action_key, current_time)
            dialog.destroy()

        def do_cancel():
            if self.timer_id:
                dialog.after_cancel(self.timer_id)
                self.timer_id = None
            self.log_message(f"Confirmation cancelled for {name}")
            dialog.destroy()

        confirm_btn = ttk.Button(btn_frame, text="Confirm", command=do_confirm, width=12)
        confirm_btn.pack(side=tk.LEFT, padx=10)
        cancel_btn = ttk.Button(btn_frame, text="Cancel", command=do_cancel, width=12)
        cancel_btn.pack(side=tk.LEFT, padx=10)

        self.update_countdown(dialog, staff_id, action_key, current_time)
        dialog.protocol("WM_DELETE_WINDOW", do_cancel)

    def update_countdown(self, dialog, staff_id, action_key, current_time):
        if self.countdown <= 0:
            self.perform_action(staff_id, action_key, current_time)
            dialog.destroy()
            return

        self.countdown_label.config(text=f"Auto‑confirm in {self.countdown} seconds")
        self.countdown -= 1
        self.timer_id = dialog.after(1000, self.update_countdown, dialog, staff_id, action_key, current_time)

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
        self.barcode_entry.config(state=tk.NORMAL)
        self.scan_btn.config(state=tk.NORMAL)
        self.confirm_dialog = None
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None

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

    def show_monthly_summary(self):
        win = tk.Toplevel(self.root)
        win.title("Monthly Summary")
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
            rows = get_monthly_summary(year, month)
            if not rows:
                messagebox.showinfo("No Data", "No attendance records for this month.")
                return

            sum_win = tk.Toplevel(win)
            sum_win.title(f"Summary for {year}-{month:02d}")
            sum_win.geometry("700x400")

            tree = ttk.Treeview(sum_win, columns=("ID", "Name", "Batch", "Total Hours", "Days Worked"), show="headings")
            tree.heading("ID", text="Staff ID")
            tree.heading("Name", text="Name")
            tree.heading("Batch", text="Batch")
            tree.heading("Total Hours", text="Total Hours")
            tree.heading("Days Worked", text="Days Worked")
            tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            for row in rows:
                staff_id, name, batch, total_sec, days = row
                hours = total_sec / 3600.0 if total_sec else 0
                tree.insert("", tk.END, values=(staff_id, name, batch, f"{hours:.2f}", days))

            def show_details(event):
                selected = tree.selection()
                if not selected:
                    return
                values = tree.item(selected[0])['values']
                staff_id = values[0]
                details = get_daily_details(staff_id, year, month)
                if not details:
                    return
                detail_win = tk.Toplevel(sum_win)
                detail_win.title(f"Daily details for {values[1]}")
                detail_win.geometry("500x300")
                dtree = ttk.Treeview(detail_win, columns=("Date", "Checkin", "Checkout", "Hours"), show="headings")
                dtree.heading("Date", text="Date")
                dtree.heading("Checkin", text="Checkin")
                dtree.heading("Checkout", text="Checkout")
                dtree.heading("Hours", text="Hours")
                dtree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
                for d in details:
                    date, checkin, checkout, hours = d
                    dtree.insert("", tk.END, values=(date, checkin, checkout, f"{hours:.2f}"))
                ttk.Button(detail_win, text="Close", command=detail_win.destroy).pack(pady=5)

            tree.bind("<Double-1>", show_details)

            def export_csv():
                import csv
                from tkinter import filedialog
                file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                         filetypes=[("CSV files", "*.csv")])
                if not file_path:
                    return
                with open(file_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Staff ID", "Name", "Batch", "Total Hours", "Days Worked"])
                    for row in rows:
                        staff_id, name, batch, total_sec, days = row
                        hours = total_sec / 3600.0 if total_sec else 0
                        writer.writerow([staff_id, name, batch, f"{hours:.2f}", days])
                messagebox.showinfo("Export", f"Summary exported to {file_path}")

            ttk.Button(sum_win, text="Export CSV", command=export_csv).pack(pady=5)

        ttk.Button(win, text="Generate", command=generate).pack(pady=20)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=5)

# ---------- Main ----------
if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
