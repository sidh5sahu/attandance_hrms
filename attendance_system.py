"""
Attendance Management System
A complete database-driven system for managing employee attendance
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
from database import AttendanceDatabase
import os

# Global database instance
db = None

class AttendanceManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Management System")
        self.root.geometry("1000x700")
        
        # Initialize database
        global db
        db = AttendanceDatabase("attendance.db")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_dashboard_tab()
        self.create_employees_tab()
        self.create_attendance_tab()
        self.create_holidays_tab()
        self.create_reports_tab()
        self.create_settings_tab()
        
    # ==================== DASHBOARD TAB ====================
    def create_dashboard_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="📊 Dashboard")
        
        # Statistics Frame
        stats_frame = tk.LabelFrame(tab, text="Statistics", font=("Arial", 12, "bold"))
        stats_frame.pack(fill="x", padx=20, pady=10)
        
        stats_grid = tk.Frame(stats_frame)
        stats_grid.pack(pady=10)
        
        # Stat cards
        self.stat_employees = self.create_stat_card(stats_grid, "Total Employees", "0", 0, 0)
        self.stat_records = self.create_stat_card(stats_grid, "Total Records", "0", 0, 1)
        self.stat_departments = self.create_stat_card(stats_grid, "Departments", "0", 1, 0)
        self.stat_today = self.create_stat_card(stats_grid, "Today's Present", "0", 1, 1)
        
        # Quick Actions
        actions_frame = tk.LabelFrame(tab, text="Quick Actions", font=("Arial", 12, "bold"))
        actions_frame.pack(fill="x", padx=20, pady=10)
        
        btn_frame = tk.Frame(actions_frame)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="📤 Upload Attendance", command=self.quick_upload, 
                  bg="#4CAF50", fg="white", font=("Arial", 11), width=20, height=2).grid(row=0, column=0, padx=10, pady=5)
        tk.Button(btn_frame, text="📋 Monthly Report", command=self.quick_monthly_report, 
                  bg="#2196F3", fg="white", font=("Arial", 11), width=20, height=2).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(btn_frame, text="🔄 Refresh Stats", command=self.refresh_dashboard, 
                  bg="#FF9800", fg="white", font=("Arial", 11), width=20, height=2).grid(row=0, column=2, padx=10, pady=5)
        
        # Load initial stats
        self.refresh_dashboard()
    
    def create_stat_card(self, parent, title, value, row, col):
        frame = tk.Frame(parent, relief="raised", borderwidth=2, bg="#f0f0f0")
        frame.grid(row=row, column=col, padx=10, pady=10, sticky="ew")
        
        tk.Label(frame, text=title, font=("Arial", 10), bg="#f0f0f0").pack(pady=(10, 5))
        value_label = tk.Label(frame, text=value, font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#2196F3")
        value_label.pack(pady=(5, 10))
        
        return value_label
    
    def refresh_dashboard(self):
        stats = db.get_statistics()
        self.stat_employees.config(text=str(stats['total_employees']))
        self.stat_records.config(text=str(stats['total_attendance_records']))
        self.stat_departments.config(text=str(stats['total_departments']))
        
        # Get today's present count
        today = datetime.now().strftime('%Y-%m-%d')
        try:
            # This is a simplified count - you may want to enhance this
            self.stat_today.config(text="N/A")
        except:
            self.stat_today.config(text="0")
        
        messagebox.showinfo("Success", "Dashboard refreshed!")
    
    def quick_upload(self):
        self.notebook.select(2)  # Switch to Attendance tab
    
    def quick_monthly_report(self):
        self.notebook.select(4)  # Switch to Reports tab
    
    # ==================== EMPLOYEES TAB ====================
    def create_employees_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="👥 Employees")
        
        # Toolbar
        toolbar = tk.Frame(tab)
        toolbar.pack(fill="x", padx=10, pady=5)
        
        tk.Button(toolbar, text="➕ Add Employee", command=self.add_employee, 
                  bg="#4CAF50", fg="white", width=15).pack(side="left", padx=5)
        tk.Button(toolbar, text="✏️ Edit", command=self.edit_employee, 
                  bg="#2196F3", fg="white", width=15).pack(side="left", padx=5)
        tk.Button(toolbar, text="🗑️ Delete", command=self.delete_employee, 
                  bg="#f44336", fg="white", width=15).pack(side="left", padx=5)
        tk.Button(toolbar, text="📤 Import from Excel", command=self.import_employees, 
                  bg="#FF9800", fg="white", width=15).pack(side="left", padx=5)
        tk.Button(toolbar, text="🔄 Refresh", command=self.refresh_employees, 
                  bg="#607D8B", fg="white", width=15).pack(side="left", padx=5)
        
        # Search
        search_frame = tk.Frame(tab)
        search_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        self.emp_search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.emp_search_var, width=30).pack(side="left", padx=5)
        tk.Button(search_frame, text="🔍 Search", command=self.search_employees).pack(side="left", padx=5)
        
        # Employee Table
        table_frame = tk.Frame(tab)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.emp_tree = ttk.Treeview(table_frame, columns=("ID", "Name", "Department"), show="headings", height=20)
        self.emp_tree.heading("ID", text="Employee ID")
        self.emp_tree.heading("Name", text="Name")
        self.emp_tree.heading("Department", text="Department")
        self.emp_tree.column("ID", width=150)
        self.emp_tree.column("Name", width=300)
        self.emp_tree.column("Department", width=200)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.emp_tree.yview)
        self.emp_tree.configure(yscrollcommand=scrollbar.set)
        self.emp_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Load employees
        self.refresh_employees()
    
    def refresh_employees(self):
        for item in self.emp_tree.get_children():
            self.emp_tree.delete(item)
        
        employees = db.get_all_employees()
        for emp in employees:
            self.emp_tree.insert("", "end", values=(emp['emp_id'], emp['name'], emp['dept']))
    
    def search_employees(self):
        search_term = self.emp_search_var.get().lower()
        for item in self.emp_tree.get_children():
            self.emp_tree.delete(item)
        
        employees = db.get_all_employees()
        for emp in employees:
            if search_term in emp['emp_id'].lower() or search_term in emp['name'].lower() or search_term in emp['dept'].lower():
                self.emp_tree.insert("", "end", values=(emp['emp_id'], emp['name'], emp['dept']))
    
    def add_employee(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Employee")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Employee ID:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        emp_id_entry = tk.Entry(dialog, width=30)
        emp_id_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Name:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        name_entry = tk.Entry(dialog, width=30)
        name_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Department:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        dept_entry = tk.Entry(dialog, width=30)
        dept_entry.grid(row=2, column=1, padx=10, pady=10)
        
        def save():
            emp_id = emp_id_entry.get().strip()
            name = name_entry.get().strip()
            dept = dept_entry.get().strip()
            
            if not emp_id or not name or not dept:
                messagebox.showerror("Error", "All fields are required")
                return
            
            if db.add_employee(emp_id, name, dept):
                messagebox.showinfo("Success", "Employee added successfully!")
                self.refresh_employees()
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Employee ID already exists")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
        tk.Button(btn_frame, text="Save", command=save, bg="green", fg="white", width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=10).pack(side="left", padx=5)
    
    def edit_employee(self):
        selected = self.emp_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select an employee")
            return
        
        values = self.emp_tree.item(selected[0])['values']
        current_id = str(values[0])
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Employee")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Employee ID:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        emp_id_label = tk.Label(dialog, text=values[0], font=("Arial", 10, "bold"))
        emp_id_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        tk.Label(dialog, text="Name:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        name_entry = tk.Entry(dialog, width=30)
        name_entry.insert(0, values[1])
        name_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Department:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        dept_entry = tk.Entry(dialog, width=30)
        dept_entry.insert(0, values[2])
        dept_entry.grid(row=2, column=1, padx=10, pady=10)
        
        def save():
            name = name_entry.get().strip()
            dept = dept_entry.get().strip()
            
            if not name or not dept:
                messagebox.showerror("Error", "All fields are required")
                return
            
            if db.update_employee(current_id, name, dept):
                messagebox.showinfo("Success", "Employee updated successfully!")
                self.refresh_employees()
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Update failed")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
        tk.Button(btn_frame, text="Save", command=save, bg="green", fg="white", width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=10).pack(side="left", padx=5)
    
    def delete_employee(self):
        selected = self.emp_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select an employee")
            return
        
        values = self.emp_tree.item(selected[0])['values']
        
        if messagebox.askyesno("Confirm Delete", f"Delete employee {values[0]} - {values[1]}?\\nThis will also delete all attendance records!"):
            if db.delete_employee(str(values[0])):
                messagebox.showinfo("Success", "Employee deleted successfully!")
                self.refresh_employees()
    
    def import_employees(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                success, errors = db.import_employees_from_excel(file_path)
                messagebox.showinfo("Import Complete", f"Imported {success} employees\\n{errors} duplicates/errors")
                self.refresh_employees()
            except Exception as e:
                messagebox.showerror("Error", f"Import failed: {str(e)}")
    
    # ==================== ATTENDANCE TAB ====================
    def create_attendance_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="📋 Attendance")
        
        # Upload Section
        upload_frame = tk.LabelFrame(tab, text="Upload Attendance Data", font=("Arial", 11, "bold"))
        upload_frame.pack(fill="x", padx=20, pady=10)
        
        format_label = tk.Label(upload_frame, 
                               text="📋 Required Format: 3 columns only\n\n"
                               "Column 1: emp_id (Employee ID)\n"
                               "Column 2: date (YYYY-MM-DD or any date format)\n"
                               "Column 3: time (HH:MM:SS format)\n\n"
                               "⚠️ All other columns will be ignored.\n"
                               "System will calculate: Punch In = First entry, Punch Out = Last entry", 
                               font=("Arial", 9), fg="#666", justify="left")
        format_label.pack(pady=10, padx=10)
        
        btn_frame = tk.Frame(upload_frame)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="📁 Choose File & Upload", command=self.upload_attendance_file, 
                  bg="#4CAF50", fg="white", font=("Arial", 11), width=25).pack(pady=5)
        
        # Manual Entry Section
        manual_frame = tk.LabelFrame(tab, text="Manual Attendance Entry", font=("Arial", 11, "bold"))
        manual_frame.pack(fill="x", padx=20, pady=10)
        
        form_frame = tk.Frame(manual_frame)
        form_frame.pack(pady=10)
        
        tk.Label(form_frame, text="Date (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.att_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form_frame, textvariable=self.att_date_var, width=20).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(form_frame, text="Employee:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.att_emp_var = tk.StringVar()
        self.att_emp_combo = ttk.Combobox(form_frame, textvariable=self.att_emp_var, width=30, state="readonly")
        self.att_emp_combo.grid(row=1, column=1, padx=5, pady=5)
        self.refresh_employee_dropdown()
        
        tk.Label(form_frame, text="Punch In (HH:MM):").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.punch_in_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.punch_in_var, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(form_frame, text="Punch Out (HH:MM):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.punch_out_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.punch_out_var, width=10).grid(row=1, column=3, padx=5, pady=5)
        
        tk.Label(form_frame, text="Status:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.att_status_var = tk.StringVar(value="P")
        status_combo = ttk.Combobox(form_frame, textvariable=self.att_status_var, width=18, state="readonly",
                                     values=["P", "A", "CL", "EL", "RH", "HD", "COFF", "DL", "WO", "H"])
        status_combo.grid(row=2, column=1, padx=5, pady=5)
        
        tk.Button(form_frame, text="💾 Save Attendance", command=self.save_manual_attendance, 
                  bg="#2196F3", fg="white", width=20).grid(row=2, column=2, columnspan=2, padx=5, pady=15)
        
        # Status legend
        legend_frame = tk.LabelFrame(tab, text="Status Codes", font=("Arial", 10, "bold"))
        legend_frame.pack(fill="x", padx=20, pady=10)
        
        legend_text = """P=Present | A=Absent | CL=Casual Leave | EL=Earned Leave | RH=Restricted Holiday | 
HD=Half Day | COFF=Compensatory Off | DL=Duty Leave | WO=Weekly Off | H=Holiday"""
        tk.Label(legend_frame, text=legend_text, font=("Arial", 9), justify="left").pack(pady=5, padx=10)
    
    def refresh_employee_dropdown(self):
        employees = db.get_all_employees()
        emp_list = [f"{emp['emp_id']} - {emp['name']}" for emp in employees]
        self.att_emp_combo['values'] = emp_list
        if emp_list:
            self.att_emp_combo.current(0)
    
    def upload_attendance_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if not file_path:
            return
        
        try:
            # Read file
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            # Process and upload
            new_records, skipped_records, month_summary = self.process_attendance_file(df)
            
            # Show detailed summary
            summary_msg = f"Upload Summary:\n\n"
            summary_msg += f"✅ New records added: {new_records}\n"
            summary_msg += f"⏭️ Duplicates skipped: {skipped_records}\n\n"
            summary_msg += f"Month-wise breakdown:\n"
            for month, count in sorted(month_summary.items()):
                summary_msg += f"  {month}: {count} records\n"
            
            messagebox.showinfo("Upload Complete", summary_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Upload failed: {str(e)}")
    
    def process_attendance_file(self, df):
        """Process attendance DataFrame - uses only 3 columns: emp_id, date, time (HH:MM:SS)
        Ignores all other columns. Calculates punch_in (first entry) and punch_out (last entry) per day."""
        
        # Validate required columns
        required_cols = ['emp_id', 'date', 'time']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            raise ValueError(f"Missing required columns: {', '.join(missing_cols)}\\n\\nRequired format: emp_id, date, time (HH:MM:SS)")
        
        # Use only the 3 required columns, ignore everything else
        df = df[required_cols].copy()
        
        # Convert date to standard format
        df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')
        
        # Convert time to datetime to enable min/max operations
        df['time_obj'] = pd.to_datetime(df['time'], errors='coerce')
        
        # Remove invalid time entries
        df = df.dropna(subset=['time_obj'])
        
        # Group by emp_id and date to calculate punch in/out
        records = []
        month_summary = {}
        
        grouped = df.groupby(['emp_id', 'date'])
        
        for (emp_id, date), group in grouped:
            # Calculate punch times from all entries
            times = group['time_obj']
            punch_in = times.min().strftime('%H:%M:%S')
            punch_out = times.max().strftime('%H:%M:%S')
            
            # Determine status based on date
            if db.is_weekend(date):
                status = 'WO'
            elif db.is_holiday(date):
                status = 'H'
            else:
                status = 'P'
            
            records.append((str(emp_id), date, status, punch_in, punch_out))
            
            # Track month-wise
            month_key = date[:7]  # YYYY-MM
            month_summary[month_key] = month_summary.get(month_key, 0) + 1
        
        # Use smart bulk upload (skips duplicates)
        new_count, skip_count = db.add_attendance_bulk_smart(records)
        
        return new_count, skip_count, month_summary
    
    def save_manual_attendance(self):
        date = self.att_date_var.get()
        emp_str = self.att_emp_var.get()
        punch_in = self.punch_in_var.get().strip() or None
        punch_out = self.punch_out_var.get().strip() or None
        status = self.att_status_var.get()
        
        if not emp_str:
            messagebox.showerror("Error", "Please select an employee")
            return
        
        emp_id = emp_str.split(" - ")[0]
        
        # Check if weekend or holiday
        if db.is_weekend(date):
            if messagebox.askyesno("Weekend Detected", f"{date} is a weekend. Mark as Weekly Off (WO)?"):
                status = 'WO'
        
        holiday = db.is_holiday(date)
        if holiday:
            if messagebox.askyesno("Holiday Detected", f"{date} is a holiday ({holiday}). Mark as Holiday (H)?"):
                status = 'H'
        
        if db.add_attendance(emp_id, date, status, punch_in, punch_out):
            messagebox.showinfo("Success", "Attendance saved successfully!")
            # Clear punch times
            self.punch_in_var.set("")
            self.punch_out_var.set("")
        else:
            messagebox.showerror("Error", "Failed to save attendance")
    
    # ==================== HOLIDAYS TAB ====================
    def create_holidays_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="🗓️ Holidays")
        
        # Upload Section
        upload_frame = tk.LabelFrame(tab, text="Upload Holiday List", font=("Arial", 11, "bold"))
        upload_frame.pack(fill="x", padx=20, pady=10)
        
        tk.Label(upload_frame, text="Upload Excel/CSV with columns: date, holiday_name, type (optional)", 
                 font=("Arial", 9), fg="#666").pack(pady=5)
        
        tk.Button(upload_frame, text="📁 Upload Holiday List", command=self.upload_holidays, 
                  bg="#4CAF50", fg="white", font=("Arial", 11), width=25).pack(pady=10)
        
        # View Holidays
        view_frame = tk.LabelFrame(tab, text="View Holidays", font=("Arial", 11, "bold"))
        view_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        toolbar = tk.Frame(view_frame)
        toolbar.pack(fill="x", padx=5, pady=5)
        
        tk.Label(toolbar, text="Year:").pack(side="left", padx=5)
        self.holiday_year_var = tk.StringVar(value=str(datetime.now().year))
        tk.Spinbox(toolbar, from_=2020, to=2030, textvariable=self.holiday_year_var, width=10).pack(side="left", padx=5)
        tk.Button(toolbar, text="🔍 Load", command=self.load_holidays, bg="#2196F3", fg="white").pack(side="left", padx=5)
        tk.Button(toolbar, text="🗑️ Delete Selected", command=self.delete_holiday, bg="#f44336", fg="white").pack(side="left", padx=5)
        
        # Holidays table
        table_frame = tk.Frame(view_frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.holiday_tree = ttk.Treeview(table_frame, columns=("ID", "Date", "Name", "Type"), show="headings", height=15)
        self.holiday_tree.heading("ID", text="ID")
        self.holiday_tree.heading("Date", text="Date")
        self.holiday_tree.heading("Name", text="Holiday Name")
        self.holiday_tree.heading("Type", text="Type")
        self.holiday_tree.column("ID", width=50)
        self.holiday_tree.column("Date", width=120)
        self.holiday_tree.column("Name", width=300)
        self.holiday_tree.column("Type", width=100)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.holiday_tree.yview)
        self.holiday_tree.configure(yscrollcommand=scrollbar.set)
        self.holiday_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        tk.Label(view_frame, text="Note: Saturdays and Sundays are automatically marked as Weekly Off (WO)", 
                 font=("Arial", 9, "italic"), fg="#666").pack(pady=5)
    
    def upload_holidays(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if not file_path:
            return
        
        try:
            success, errors = db.import_holidays_from_excel(file_path)
            messagebox.showinfo("Success", f"Imported {success} holidays\\n{errors} duplicates/errors")
            self.load_holidays()
        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {str(e)}")
    
    def load_holidays(self):
        for item in self.holiday_tree.get_children():
            self.holiday_tree.delete(item)
        
        try:
            year = int(self.holiday_year_var.get())
            holidays = db.get_holidays_by_year(year)
            for hol in holidays:
                self.holiday_tree.insert("", "end", values=(hol['id'], hol['date'], hol['holiday_name'], hol['type']))
        except:
            messagebox.showerror("Error", "Invalid year")
    
    def delete_holiday(self):
        selected = self.holiday_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a holiday")
            return
        
        hol_id = self.holiday_tree.item(selected[0])['values'][0]
        if messagebox.askyesno("Confirm", "Delete this holiday?"):
            if db.delete_holiday(hol_id):
                messagebox.showinfo("Success", "Holiday deleted")
                self.load_holidays()
    
    # ==================== REPORTS TAB ====================
    def create_reports_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="📊 Reports")
        
        # Report Type Selection
        type_frame = tk.LabelFrame(tab, text="Select Report Type", font=("Arial", 11, "bold"))
        type_frame.pack(fill="x", padx=20, pady=10)
        
        self.report_type_var = tk.StringVar(value="employee")
        tk.Radiobutton(type_frame, text="📄 Employee Report", variable=self.report_type_var, value="employee", 
                       font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        tk.Radiobutton(type_frame, text="🏢 Department Report", variable=self.report_type_var, value="department", 
                       font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        tk.Radiobutton(type_frame, text="📅 Monthly Report", variable=self.report_type_var, value="monthly", 
                       font=("Arial", 10)).pack(anchor="w", padx=20, pady=5)
        
        # Parameters
        params_frame = tk.LabelFrame(tab, text="Report Parameters", font=("Arial", 11, "bold"))
        params_frame.pack(fill="x", padx=20, pady=10)
        
        grid_frame = tk.Frame(params_frame)
        grid_frame.pack(pady=10)
        
        tk.Label(grid_frame, text="Employee:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.report_emp_var = tk.StringVar()
        self.report_emp_combo = ttk.Combobox(grid_frame, textvariable=self.report_emp_var, width=30, state="readonly")
        self.report_emp_combo.grid(row=0, column=1, padx=5, pady=5)
        self.report_emp_combo['values'] = [f"{emp['emp_id']} - {emp['name']}" for emp in db.get_all_employees()]
        
        tk.Label(grid_frame, text="Department:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.report_dept_var = tk.StringVar()
        dept_combo = ttk.Combobox(grid_frame, textvariable=self.report_dept_var, width=30, state="readonly")
        dept_combo.grid(row=1, column=1, padx=5, pady=5)
        dept_combo['values'] = db.get_all_departments()
        
        tk.Label(grid_frame, text="Year:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.report_year_var = tk.StringVar(value=str(datetime.now().year))
        tk.Spinbox(grid_frame, from_=2020, to=2030, textvariable=self.report_year_var, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(grid_frame, text="Month:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.report_month_var = tk.StringVar(value=str(datetime.now().month))
        tk.Spinbox(grid_frame, from_=1, to=12, textvariable=self.report_month_var, width=10).grid(row=1, column=3, padx=5, pady=5)
        
        tk.Label(grid_frame, text="Start Date:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.report_start_var = tk.StringVar()
        tk.Entry(grid_frame, textvariable=self.report_start_var, width=32).grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(grid_frame, text="End Date:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.report_end_var = tk.StringVar()
        tk.Entry(grid_frame, textvariable=self.report_end_var, width=12).grid(row=2, column=3, padx=5, pady=5)
        
        # Generate Button
        tk.Button(params_frame, text="📊 Generate Report", command=self.generate_report, 
                  bg="#673AB7", fg="white", font=("Arial", 12, "bold"), width=25, height=2).pack(pady=15)
    
    def generate_report(self):
        report_type = self.report_type_var.get()
        
        try:
            if report_type == "employee":
                self.generate_employee_report()
            elif report_type == "department":
                self.generate_department_report()
            elif report_type == "monthly":
                self.generate_monthly_report()
        except Exception as e:
            messagebox.showerror("Error", f"Report generation failed: {str(e)}")
    
    def generate_employee_report(self):
        emp_str = self.report_emp_var.get()
        if not emp_str:
            messagebox.showerror("Error", "Please select an employee")
            return
        
        emp_id = emp_str.split(" - ")[0]
        start_date = self.report_start_var.get() or None
        end_date = self.report_end_var.get() or None
        
        # Get data
        df = db.get_employee_attendance(emp_id, start_date, end_date)
        
        if df.empty:
            messagebox.showinfo("No Data", "No attendance records found")
            return
        
        # Save to Excel
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                  filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Report saved to {save_path}")
    
    def generate_department_report(self):
        dept = self.report_dept_var.get()
        if not dept:
            messagebox.showerror("Error", "Please select a department")
            return
        
        start_date = self.report_start_var.get() or None
        end_date = self.report_end_var.get() or None
        
        df = db.get_dept_attendance(dept, start_date, end_date)
        
        if df.empty:
            messagebox.showinfo("No Data", "No attendance records found")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                  filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Report saved to {save_path}")
    
    def generate_monthly_report(self):
        try:
            year = int(self.report_year_var.get())
            month = int(self.report_month_var.get())
        except:
            messagebox.showerror("Error", "Invalid year or month")
            return
        
        df = db.get_month_attendance(year, month)
        
        if df.empty:
            messagebox.showinfo("No Data", "No attendance records found")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                  filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Report saved to {save_path}")
    
    # ==================== SETTINGS TAB ====================
    def create_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="⚙️ Settings")
        
        # Database Management
        db_frame = tk.LabelFrame(tab, text="Database Management", font=("Arial", 11, "bold"))
        db_frame.pack(fill="x", padx=20, pady=10)
        
        btn_frame = tk.Frame(db_frame)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="💾 Backup Database", command=self.backup_db, 
                  bg="#9C27B0", fg="white", width=20).grid(row=0, column=0, padx=10, pady=5)
        tk.Button(btn_frame, text="📤 Export All to Excel", command=self.export_all, 
                  bg="#FF9800", fg="white", width=20).grid(row=0, column=1, padx=10, pady=5)
        
        # Info
        info_frame = tk.LabelFrame(tab, text="System Information", font=("Arial", 11, "bold"))
        info_frame.pack(fill="x", padx=20, pady=10)
        
        self.info_label = tk.Label(info_frame, text="", font=("Arial", 10), justify="left")
        self.info_label.pack(pady=10, padx=10)
        self.update_system_info()
    
    def backup_db(self):
        try:
            backup_path = db.backup_database()
            messagebox.showinfo("Success", f"Database backed up to:\\n{backup_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Backup failed: {str(e)}")
    
    def export_all(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                  filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                db.export_to_excel(save_path)
                messagebox.showinfo("Success", f"Data exported to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def update_system_info(self):
        stats = db.get_statistics()
        info_text = f"""Database: attendance.db
Total Employees: {stats['total_employees']}
Total Attendance Records: {stats['total_attendance_records']}
Departments: {stats['total_departments']}
Date Range: {stats['earliest_attendance'] or 'N/A'} to {stats['latest_attendance'] or 'N/A'}

Leave Types:
  P - Present
  A - Absent
  CL - Casual Leave
  EL - Earned Leave
  RH - Restricted Holiday
  HD - Half Day Leave
  COFF - Compensatory Off
  DL - Duty Leave
  WO - Weekly Off (Saturday/Sunday)
  H - Holiday
"""
        self.info_label.config(text=info_text)

def main():
    root = tk.Tk()
    app = AttendanceManagementSystem(root)
    root.mainloop()

if __name__ == "__main__":
    main()
