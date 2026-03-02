import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime
from openpyxl.styles import Font
import os
import shutil
from database import AttendanceDatabase

# Global variables
emp_file = ""
att_file = ""
employee_df = None
db = None  # Database instance

def load_employee_file():
    """Load employee file and return DataFrame"""
    global emp_file, employee_df
    if emp_file and os.path.exists(emp_file):
        try:
            employee_df = pd.read_excel(emp_file)
            if not all(col in employee_df.columns for col in ['emp_id', 'name', 'Dept']):
                messagebox.showerror("Error", "Employee file must have columns: emp_id, name, Dept")
                employee_df = None
                return None
            return employee_df
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employee file: {str(e)}")
            return None
    return None

def save_employee_file():
    """Save employee DataFrame back to file"""
    global emp_file, employee_df
    if emp_file and employee_df is not None:
        try:
            # Create backup
            backup_path = emp_file + ".backup"
            if os.path.exists(emp_file):
                shutil.copy(emp_file, backup_path)
            employee_df.to_excel(emp_file, index=False)
            messagebox.showinfo("Success", "Employee data saved successfully!")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {str(e)}")
            return False
    else:
        messagebox.showerror("Error", "No employee file loaded")
        return False

def browse_emp_file():
    global emp_file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        emp_file = file_path
        emp_path.set(file_path)
        load_employee_file()
        refresh_employee_list()
        refresh_employee_dropdown()

def browse_att_file():
    global att_file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        att_file = file_path
        att_path.set(file_path)

def generate():
    global emp_file, att_file
    start_date_str = start_date_var.get()
    end_date_str = end_date_var.get()

    if emp_file == "" or att_file == "":
        messagebox.showerror("Error", "Select both files")
        return

    # Load files
    emp = pd.read_excel(emp_file)
    att = pd.read_excel(att_file)

    # Columns must be: emp_id, name, Dept
    emp = emp[['emp_id', 'name', 'Dept']]
    att['date'] = pd.to_datetime(att['date'])

    # Determine Date Range
    if start_date_str and end_date_str:
        try:
            start = pd.to_datetime(start_date_str)
            end = pd.to_datetime(end_date_str)
        except Exception:
            messagebox.showerror("Error", "Invalid Date Format. Use YYYY-MM-DD")
            return
    else:
        # Default to full month of first record
        month = att['date'].dt.month.iloc[0]
        year = att['date'].dt.year.iloc[0]
        start = pd.Timestamp(year, month, 1)
        end = start + pd.offsets.MonthEnd(1)

    all_dates = pd.date_range(start, end)

    base = emp.copy()

    for d in all_dates:
        base[d.strftime("%Y-%m-%d")] = ""

    # Holiday rule
    for d in all_dates:
        if d.weekday() in (5, 6):
            base[d.strftime("%Y-%m-%d")] = "WO"
        elif d.strftime("%Y-%m-%d") == "2025-12-25":
            base[d.strftime("%Y-%m-%d")] = "H"

    # Mark Present
    for _, r in att.iterrows():
        d_timestamp = r['date']
        if start <= d_timestamp <= end:
            d = d_timestamp.strftime("%Y-%m-%d")
            if d in base.columns:
                base.loc[base.emp_id == r.emp_id, d] = "P"

    # Mark Absents
    for d in all_dates:
        col = d.strftime("%Y-%m-%d")
        base.loc[base[col] == "", col] = "A"

    try:
        # Save dept-wise
        save = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save == "":
            return

        with pd.ExcelWriter(save, engine="openpyxl") as w:
            for dept, g in base.groupby("Dept"):
                g = g.copy()
                date_cols = [c for c in g.columns if c not in ["emp_id", "name", "Dept"]]
                g["Total_Present"] = (g[date_cols] == "P").sum(axis=1)
                g["Total_Absent"] = (g[date_cols] == "A").sum(axis=1)
                g["Total_Holidays"] = (g[date_cols] == "H").sum(axis=1)
                g["Total_Weekly_Off"] = (g[date_cols] == "WO").sum(axis=1)

                # Create a row for Day Names (e.g., Monday, Tuesday)
                day_row_values = []
                for c in g.columns:
                    if c in date_cols:
                        day_name = datetime.strptime(c, "%Y-%m-%d").strftime("%A")
                        day_row_values.append(day_name)
                    else:
                        day_row_values.append("")

                # Create a DataFrame for the day row
                day_row_df = pd.DataFrame([day_row_values], columns=g.columns)

                # Concatenate the day row at the top of the data
                final_df = pd.concat([day_row_df, g], ignore_index=True)

                sheet_name = str(dept)[:31]
                final_df.to_excel(w, sheet_name=sheet_name, index=False)

                # Apply Red Color to Weekly Offs
                ws = w.sheets[sheet_name]
                red_font = Font(color="FF0000")

                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value == "WO":
                            cell.font = red_font

        messagebox.showinfo("Success", "Attendance file generated successfully")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# ================ EMPLOYEE MANAGEMENT FUNCTIONS ================

def refresh_employee_list():
    """Refresh the employee treeview"""
    global employee_df
    for item in emp_tree.get_children():
        emp_tree.delete(item)
    
    if employee_df is not None:
        for idx, row in employee_df.iterrows():
            emp_tree.insert("", "end", values=(row['emp_id'], row['name'], row['Dept']))

def add_employee():
    """Add a new employee"""
    global employee_df
    
    if employee_df is None:
        messagebox.showerror("Error", "Please load an employee file first")
        return
    
    # Create add employee dialog
    dialog = tk.Toplevel(app)
    dialog.title("Add Employee")
    dialog.geometry("350x200")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()
    
    # Center the dialog
    dialog.geometry("+%d+%d" % (app.winfo_x() + 100, app.winfo_y() + 100))
    
    tk.Label(dialog, text="Employee ID:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    emp_id_entry = tk.Entry(dialog, width=30)
    emp_id_entry.grid(row=0, column=1, padx=10, pady=10)
    
    tk.Label(dialog, text="Name:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    name_entry = tk.Entry(dialog, width=30)
    name_entry.grid(row=1, column=1, padx=10, pady=10)
    
    tk.Label(dialog, text="Department:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    dept_entry = tk.Entry(dialog, width=30)
    dept_entry.grid(row=2, column=1, padx=10, pady=10)
    
    def save_new_employee():
        global employee_df
        emp_id = emp_id_entry.get().strip()
        name = name_entry.get().strip()
        dept = dept_entry.get().strip()
        
        if not emp_id or not name or not dept:
            messagebox.showerror("Error", "All fields are required")
            return
        
        # Check for duplicate emp_id
        if emp_id in employee_df['emp_id'].astype(str).values:
            messagebox.showerror("Error", "Employee ID already exists")
            return
        
        # Add new employee
        new_row = pd.DataFrame({'emp_id': [emp_id], 'name': [name], 'Dept': [dept]})
        employee_df = pd.concat([employee_df, new_row], ignore_index=True)
        
        refresh_employee_list()
        refresh_employee_dropdown()
        dialog.destroy()
        messagebox.showinfo("Success", "Employee added successfully!")
    
    btn_frame = tk.Frame(dialog)
    btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
    tk.Button(btn_frame, text="Save", command=save_new_employee, bg="green", fg="white", width=10).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=10).pack(side="left", padx=5)

def edit_employee():
    """Edit selected employee"""
    global employee_df
    
    selected = emp_tree.selection()
    if not selected:
        messagebox.showerror("Error", "Please select an employee to edit")
        return
    
    item = emp_tree.item(selected[0])
    values = item['values']
    current_id = str(values[0])
    
    # Create edit dialog
    dialog = tk.Toplevel(app)
    dialog.title("Edit Employee")
    dialog.geometry("350x200")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()
    
    dialog.geometry("+%d+%d" % (app.winfo_x() + 100, app.winfo_y() + 100))
    
    tk.Label(dialog, text="Employee ID:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    emp_id_entry = tk.Entry(dialog, width=30)
    emp_id_entry.insert(0, values[0])
    emp_id_entry.grid(row=0, column=1, padx=10, pady=10)
    
    tk.Label(dialog, text="Name:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    name_entry = tk.Entry(dialog, width=30)
    name_entry.insert(0, values[1])
    name_entry.grid(row=1, column=1, padx=10, pady=10)
    
    tk.Label(dialog, text="Department:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    dept_entry = tk.Entry(dialog, width=30)
    dept_entry.insert(0, values[2])
    dept_entry.grid(row=2, column=1, padx=10, pady=10)
    
    def save_edited_employee():
        global employee_df
        new_id = emp_id_entry.get().strip()
        new_name = name_entry.get().strip()
        new_dept = dept_entry.get().strip()
        
        if not new_id or not new_name or not new_dept:
            messagebox.showerror("Error", "All fields are required")
            return
        
        # Check for duplicate emp_id (if changed)
        if new_id != current_id and new_id in employee_df['emp_id'].astype(str).values:
            messagebox.showerror("Error", "Employee ID already exists")
            return
        
        # Update employee
        mask = employee_df['emp_id'].astype(str) == current_id
        employee_df.loc[mask, 'emp_id'] = new_id
        employee_df.loc[mask, 'name'] = new_name
        employee_df.loc[mask, 'Dept'] = new_dept
        
        refresh_employee_list()
        refresh_employee_dropdown()
        dialog.destroy()
        messagebox.showinfo("Success", "Employee updated successfully!")
    
    btn_frame = tk.Frame(dialog)
    btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
    tk.Button(btn_frame, text="Save", command=save_edited_employee, bg="green", fg="white", width=10).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=10).pack(side="left", padx=5)

def delete_employee():
    """Delete selected employee"""
    global employee_df
    
    selected = emp_tree.selection()
    if not selected:
        messagebox.showerror("Error", "Please select an employee to delete")
        return
    
    item = emp_tree.item(selected[0])
    values = item['values']
    
    if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete employee:\n\nID: {values[0]}\nName: {values[1]}\nDepartment: {values[2]}"):
        mask = employee_df['emp_id'].astype(str) == str(values[0])
        employee_df = employee_df[~mask].reset_index(drop=True)
        refresh_employee_list()
        refresh_employee_dropdown()
        messagebox.showinfo("Success", "Employee deleted successfully!")

# ================ INDIVIDUAL ATTENDANCE VIEW FUNCTIONS ================

def refresh_employee_dropdown():
    """Refresh the employee dropdown in view tab"""
    global employee_df
    if employee_df is not None:
        emp_list = [f"{row['emp_id']} - {row['name']}" for _, row in employee_df.iterrows()]
        emp_select_combo['values'] = emp_list
        if emp_list:
            emp_select_combo.current(0)
    else:
        emp_select_combo['values'] = []

def view_individual_attendance():
    """View attendance for selected employee"""
    global att_file, employee_df
    
    if not att_file:
        messagebox.showerror("Error", "Please load an attendance file first (from Generate tab)")
        return
    
    selected_emp = emp_select_var.get()
    if not selected_emp:
        messagebox.showerror("Error", "Please select an employee")
        return
    
    # Extract emp_id from selection
    emp_id = selected_emp.split(" - ")[0].strip()
    
    start_str = view_start_date_var.get()
    end_str = view_end_date_var.get()
    
    try:
        att = pd.read_excel(att_file)
        att['date'] = pd.to_datetime(att['date'])
        
        # Filter by employee
        emp_att = att[att['emp_id'].astype(str) == str(emp_id)]
        
        # Determine date range
        if start_str and end_str:
            try:
                start = pd.to_datetime(start_str)
                end = pd.to_datetime(end_str)
            except:
                messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
                return
        else:
            if len(emp_att) > 0:
                start = emp_att['date'].min()
                end = emp_att['date'].max()
            else:
                messagebox.showinfo("Info", "No attendance records found for this employee")
                return
        
        # Generate all dates in range
        all_dates = pd.date_range(start, end)
        
        # Clear existing data
        for item in att_tree.get_children():
            att_tree.delete(item)
        
        # Track counts
        total_p = 0
        total_a = 0
        total_wo = 0
        total_h = 0
        
        # Populate attendance
        present_dates = set(emp_att['date'].dt.strftime("%Y-%m-%d").values)
        
        for d in all_dates:
            d_str = d.strftime("%Y-%m-%d")
            day_name = d.strftime("%A")
            
            if d.weekday() in (5, 6):
                status = "WO"
                total_wo += 1
            elif d_str == "2025-12-25":
                status = "H"
                total_h += 1
            elif d_str in present_dates:
                status = "P"
                total_p += 1
            else:
                status = "A"
                total_a += 1
            
            att_tree.insert("", "end", values=(d_str, day_name, status))
        
        # Update summary
        summary_label.config(text=f"Present: {total_p} | Absent: {total_a} | Weekly Off: {total_wo} | Holidays: {total_h}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load attendance: {str(e)}")

# ================ GUI SETUP ================

app = tk.Tk()
app.title("HR Attendance Generator")
app.geometry("700x550")

# Create Notebook for tabs
notebook = ttk.Notebook(app)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# ================ TAB 1: GENERATE ATTENDANCE ================
tab_generate = ttk.Frame(notebook)
notebook.add(tab_generate, text="Generate Attendance")

emp_path = tk.StringVar()
att_path = tk.StringVar()
start_date_var = tk.StringVar()
end_date_var = tk.StringVar()

gen_frame = tk.Frame(tab_generate)
gen_frame.pack(pady=20)

tk.Label(gen_frame, text="Employee Sheet:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
tk.Entry(gen_frame, textvariable=emp_path, width=45).grid(row=0, column=1, padx=5, pady=10)
tk.Button(gen_frame, text="Browse", command=browse_emp_file).grid(row=0, column=2, padx=5, pady=10)

tk.Label(gen_frame, text="Attendance Sheet:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
tk.Entry(gen_frame, textvariable=att_path, width=45).grid(row=1, column=1, padx=5, pady=10)
tk.Button(gen_frame, text="Browse", command=browse_att_file).grid(row=1, column=2, padx=5, pady=10)

tk.Label(gen_frame, text="Start Date (YYYY-MM-DD):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
tk.Entry(gen_frame, textvariable=start_date_var, width=45).grid(row=2, column=1, padx=5, pady=10)

tk.Label(gen_frame, text="End Date (YYYY-MM-DD):").grid(row=3, column=0, padx=5, pady=10, sticky="e")
tk.Entry(gen_frame, textvariable=end_date_var, width=45).grid(row=3, column=1, padx=5, pady=10)

tk.Button(gen_frame, text="Generate Attendance Report", command=generate, bg="green", fg="white", 
          font=("Arial", 11, "bold"), padx=20, pady=5).grid(row=4, column=0, columnspan=3, pady=30)

# ================ TAB 2: EMPLOYEE MANAGEMENT ================
tab_employees = ttk.Frame(notebook)
notebook.add(tab_employees, text="Employee Management")

# Employee Treeview
tree_frame = tk.Frame(tab_employees)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

emp_tree = ttk.Treeview(tree_frame, columns=("ID", "Name", "Department"), show="headings", height=15)
emp_tree.heading("ID", text="Employee ID")
emp_tree.heading("Name", text="Name")
emp_tree.heading("Department", text="Department")
emp_tree.column("ID", width=120)
emp_tree.column("Name", width=250)
emp_tree.column("Department", width=200)

# Scrollbar for treeview
scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=emp_tree.yview)
emp_tree.configure(yscrollcommand=scrollbar.set)
emp_tree.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Buttons for employee management
btn_frame = tk.Frame(tab_employees)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Add Employee", command=add_employee, bg="#4CAF50", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Edit Employee", command=edit_employee, bg="#2196F3", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Delete Employee", command=delete_employee, bg="#f44336", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Save to File", command=save_employee_file, bg="#FF9800", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)

# ================ TAB 3: VIEW INDIVIDUAL ATTENDANCE ================
tab_view = ttk.Frame(notebook)
notebook.add(tab_view, text="Individual Attendance")

# Selection Frame
select_frame = tk.Frame(tab_view)
select_frame.pack(pady=15)

tk.Label(select_frame, text="Select Employee:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
emp_select_var = tk.StringVar()
emp_select_combo = ttk.Combobox(select_frame, textvariable=emp_select_var, width=40, state="readonly")
emp_select_combo.grid(row=0, column=1, padx=5, pady=5)

view_start_date_var = tk.StringVar()
view_end_date_var = tk.StringVar()

tk.Label(select_frame, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Entry(select_frame, textvariable=view_start_date_var, width=42).grid(row=1, column=1, padx=5, pady=5)

tk.Label(select_frame, text="End Date (YYYY-MM-DD):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
tk.Entry(select_frame, textvariable=view_end_date_var, width=42).grid(row=2, column=1, padx=5, pady=5)

tk.Button(select_frame, text="View Attendance", command=view_individual_attendance, bg="#673AB7", fg="white",
          font=("Arial", 10, "bold")).grid(row=3, column=0, columnspan=2, pady=15)

# Attendance Treeview
att_tree_frame = tk.Frame(tab_view)
att_tree_frame.pack(fill="both", expand=True, padx=10)

att_tree = ttk.Treeview(att_tree_frame, columns=("Date", "Day", "Status"), show="headings", height=12)
att_tree.heading("Date", text="Date")
att_tree.heading("Day", text="Day")
att_tree.heading("Status", text="Status")
att_tree.column("Date", width=150)
att_tree.column("Day", width=150)
att_tree.column("Status", width=100)

att_scrollbar = ttk.Scrollbar(att_tree_frame, orient="vertical", command=att_tree.yview)
att_tree.configure(yscrollcommand=att_scrollbar.set)
att_tree.pack(side="left", fill="both", expand=True)
att_scrollbar.pack(side="right", fill="y")

# Summary Label
summary_label = tk.Label(tab_view, text="", font=("Arial", 11, "bold"), fg="#333")
summary_label.pack(pady=10)

# ================ TAB 4: DATABASE MANAGEMENT ================
tab_database = ttk.Frame(notebook)
notebook.add(tab_database, text="Database Management")

# Database Info Frame
db_info_frame = tk.LabelFrame(tab_database, text="Database Information", font=("Arial", 10, "bold"))
db_info_frame.pack(fill="x", padx=10, pady=10)

db_status_label = tk.Label(db_info_frame, text="Database: Not Connected", font=("Arial", 10))
db_status_label.pack(pady=5)

db_stats_label = tk.Label(db_info_frame, text="", font=("Arial", 9), justify="left")
db_stats_label.pack(pady=5)

def init_database():
    """Initialize database connection"""
    global db
    try:
        db = AttendanceDatabase("attendance.db")
        update_db_status()
        messagebox.showinfo("Success", "Database connected successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to connect to database: {str(e)}")

def update_db_status():
    """Update database status display"""
    global db
    if db:
        try:
            stats = db.get_statistics()
            db_status_label.config(text="Database: Connected (attendance.db)", fg="green")
            stats_text = f"""Total Employees: {stats['total_employees']}
Total Attendance Records: {stats['total_attendance_records']}
Departments: {stats['total_departments']}
Date Range: {stats['earliest_attendance'] or 'N/A'} to {stats['latest_attendance'] or 'N/A'}"""
            db_stats_label.config(text=stats_text, fg="#333")
        except Exception as e:
            db_status_label.config(text=f"Database: Error - {str(e)}", fg="red")
    else:
        db_status_label.config(text="Database: Not Connected", fg="red")

def import_employees_to_db():
    """Import employees from Excel to database"""
    global db
    if not db:
        messagebox.showerror("Error", "Please connect to database first")
        return
    
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            success, errors = db.import_employees_from_excel(file_path)
            messagebox.showinfo("Import Complete", f"Imported {success} employees successfully\n{errors} errors/duplicates")
            update_db_status()
        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {str(e)}")

def import_attendance_to_db():
    """Import attendance from Excel to database"""
    global db
    if not db:
        messagebox.showerror("Error", "Please connect to database first")
        return
    
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            success, errors = db.import_attendance_from_excel(file_path)
            messagebox.showinfo("Import Complete", f"Imported {success} attendance records successfully\n{errors} errors")
            update_db_status()
        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {str(e)}")

def export_db_to_excel():
    """Export database to Excel"""
    global db
    if not db:
        messagebox.showerror("Error", "Please connect to database first")
        return
    
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        try:
            db.export_to_excel(save_path)
            messagebox.showinfo("Success", f"Data exported to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

def backup_database():
    """Create database backup"""
    global db
    if not db:
        messagebox.showerror("Error", "Please connect to database first")
        return
    
    try:
        backup_path = db.backup_database()
        messagebox.showinfo("Success", f"Database backed up to {backup_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Backup failed: {str(e)}")

# Database Action Buttons
db_btn_frame = tk.Frame(tab_database)
db_btn_frame.pack(pady=10)

tk.Button(db_btn_frame, text="Connect to Database", command=init_database, bg="#4CAF50", fg="white", 
          width=20, font=("Arial", 10)).grid(row=0, column=0, padx=5, pady=5)
tk.Button(db_btn_frame, text="Import Employees", command=import_employees_to_db, bg="#2196F3", fg="white", 
          width=20, font=("Arial", 10)).grid(row=0, column=1, padx=5, pady=5)
tk.Button(db_btn_frame, text="Import Attendance", command=import_attendance_to_db, bg="#2196F3", fg="white", 
          width=20, font=("Arial", 10)).grid(row=1, column=0, padx=5, pady=5)
tk.Button(db_btn_frame, text="Export to Excel", command=export_db_to_excel, bg="#FF9800", fg="white", 
          width=20, font=("Arial", 10)).grid(row=1, column=1, padx=5, pady=5)
tk.Button(db_btn_frame, text="Backup Database", command=backup_database, bg="#9C27B0", fg="white", 
          width=20, font=("Arial", 10)).grid(row=2, column=0, padx=5, pady=5)
tk.Button(db_btn_frame, text="Refresh Status", command=update_db_status, bg="#607D8B", fg="white", 
          width=20, font=("Arial", 10)).grid(row=2, column=1, padx=5, pady=5)

# Query Frame
query_frame = tk.LabelFrame(tab_database, text="Query Attendance", font=("Arial", 10, "bold"))
query_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Query Type Selection
query_type_frame = tk.Frame(query_frame)
query_type_frame.pack(pady=10)

tk.Label(query_type_frame, text="Query Type:", font=("Arial", 10)).pack(side="left", padx=5)
query_type_var = tk.StringVar(value="employee")
tk.Radiobutton(query_type_frame, text="By Employee", variable=query_type_var, value="employee", 
               font=("Arial", 9)).pack(side="left", padx=5)
tk.Radiobutton(query_type_frame, text="By Month", variable=query_type_var, value="month", 
               font=("Arial", 9)).pack(side="left", padx=5)
tk.Radiobutton(query_type_frame, text="By Department", variable=query_type_var, value="dept", 
               font=("Arial", 9)).pack(side="left", padx=5)

# Query Parameters Frame
query_params_frame = tk.Frame(query_frame)
query_params_frame.pack(pady=10)

tk.Label(query_params_frame, text="Employee ID:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
query_emp_id = tk.Entry(query_params_frame, width=20)
query_emp_id.grid(row=0, column=1, padx=5, pady=5)

tk.Label(query_params_frame, text="Department:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
query_dept = tk.Entry(query_params_frame, width=20)
query_dept.grid(row=1, column=1, padx=5, pady=5)

tk.Label(query_params_frame, text="Year:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
query_year = tk.Entry(query_params_frame, width=10)
query_year.insert(0, str(datetime.now().year))
query_year.grid(row=0, column=3, padx=5, pady=5)

tk.Label(query_params_frame, text="Month:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
query_month = tk.Entry(query_params_frame, width=10)
query_month.insert(0, str(datetime.now().month))
query_month.grid(row=1, column=3, padx=5, pady=5)

tk.Label(query_params_frame, text="Start Date:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
query_start = tk.Entry(query_params_frame, width=20)
query_start.grid(row=2, column=1, padx=5, pady=5)

tk.Label(query_params_frame, text="End Date:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
query_end = tk.Entry(query_params_frame, width=20)
query_end.grid(row=2, column=3, padx=5, pady=5)

def execute_query():
    """Execute database query based on selected type"""
    global db
    if not db:
        messagebox.showerror("Error", "Please connect to database first")
        return
    
    query_type = query_type_var.get()
    
    try:
        result_df = None
        
        if query_type == "employee":
            emp_id = query_emp_id.get().strip()
            if not emp_id:
                messagebox.showerror("Error", "Please enter Employee ID")
                return
            
            start = query_start.get().strip() or None
            end = query_end.get().strip() or None
            result_df = db.get_employee_attendance(emp_id, start, end)
            
            if result_df.empty:
                messagebox.showinfo("No Data", f"No attendance records found for employee {emp_id}")
                return
            
            # Show summary
            summary = db.get_attendance_summary_by_employee(emp_id, start, end)
            summary_text = f"Employee {emp_id} - Present: {summary['P']}, Absent: {summary['A']}, Weekly Off: {summary['WO']}, Holidays: {summary['H']}"
            
        elif query_type == "month":
            try:
                year = int(query_year.get())
                month = int(query_month.get())
            except ValueError:
                messagebox.showerror("Error", "Please enter valid year and month")
                return
            
            result_df = db.get_month_attendance(year, month)
            if result_df.empty:
                messagebox.showinfo("No Data", f"No attendance records found for {year}-{month:02d}")
                return
            
            summary_text = f"Attendance for {year}-{month:02d} - Total Records: {len(result_df)}"
            
        elif query_type == "dept":
            dept = query_dept.get().strip()
            if not dept:
                messagebox.showerror("Error", "Please enter Department")
                return
            
            start = query_start.get().strip() or None
            end = query_end.get().strip() or None
            result_df = db.get_dept_attendance(dept, start, end)
            
            if result_df.empty:
                messagebox.showinfo("No Data", f"No attendance records found for department {dept}")
                return
            
            summary_text = f"Department {dept} - Total Records: {len(result_df)}"
        
        # Display result in a new window
        display_query_result(result_df, summary_text)
        
    except Exception as e:
        messagebox.showerror("Error", f"Query failed: {str(e)}")

def display_query_result(df, summary_text):
    """Display query results in a new window"""
    result_window = tk.Toplevel(app)
    result_window.title("Query Results")
    result_window.geometry("800x500")
    
    # Summary label
    tk.Label(result_window, text=summary_text, font=("Arial", 11, "bold"), fg="#333").pack(pady=10)
    
    # Create treeview
    tree_frame = tk.Frame(result_window)
    tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Create scrollbars
    vsb = ttk.Scrollbar(tree_frame, orient="vertical")
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
    
    # Create treeview
    tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings", 
                        yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)
    
    # Add column headings
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    
    # Add data
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))
    
    # Grid layout
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
    
    # Export button
    def export_result():
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Results exported to {save_path}")
    
    tk.Button(result_window, text="Export to Excel", command=export_result, bg="#4CAF50", fg="white", 
              width=20, font=("Arial", 10)).pack(pady=10)

tk.Button(query_frame, text="Execute Query", command=execute_query, bg="#673AB7", fg="white", 
          font=("Arial", 11, "bold"), width=20).pack(pady=15)

app.mainloop()
