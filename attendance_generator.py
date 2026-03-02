import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment
import os
import shutil
import csv
import io
import json
from tkcalendar import DateEntry
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Global variables
emp_file = ""
att_file = ""
employee_df = None
attendance_df = None  # In-memory attendance data

# Persistent edits file
EDITS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attendance_edits.json")
EDITOR_PIN = "1111"

# All status/leave types
ALL_STATUS_OPTIONS = ['P', 'A', 'EL', 'CL', 'COFF', 'DL', 'ML', 'HD', 'CC', 'MaL', 'RH', 'WO', 'H']
LEAVE_TYPES = ['EL', 'CL', 'COFF', 'DL', 'ML', 'HD', 'CC', 'MaL', 'RH']

def load_saved_edits():
    """Load saved attendance edits from JSON file"""
    if os.path.exists(EDITS_FILE):
        try:
            with open(EDITS_FILE, 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_edits_to_file(edits):
    """Save attendance edits to JSON file"""
    try:
        with open(EDITS_FILE, 'w') as f:
            json.dump(edits, f, indent=2)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save edits: {str(e)}")
        return False

def get_edited_status(emp_id, date_str):
    """Get saved edit for an emp_id + date, or None if not edited"""
    edits = load_saved_edits()
    key = f"{emp_id}_{date_str}"
    return edits.get(key, None)

def apply_edits_to_base(base, date_cols):
    """Apply saved edits to a base attendance DataFrame"""
    edits = load_saved_edits()
    for key, status in edits.items():
        parts = key.rsplit('_', 1)
        if len(parts) == 2:
            emp_id, date_str = parts
            if date_str in date_cols:
                mask = base['emp_id'].astype(str) == str(emp_id)
                if mask.any():
                    base.loc[mask, date_str] = status
    return base


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

def parse_dat_file(filepath):
    """Parse a .dat attendance file (biometric device export).
    Auto-detects delimiter and extracts emp_id, date, time columns.
    Returns a DataFrame with columns: emp_id, date, time
    """
    rows = []
    detected_delimiter = None
    
    with open(filepath, 'r', errors='replace') as f:
        raw_lines = f.readlines()
    
    if not raw_lines:
        raise ValueError("File is empty")
    
    sample_lines = [l.strip() for l in raw_lines[:20] if l.strip()]
    
    for delim in ['\t', ',', ';', ' ']:
        counts = [len(l.split(delim)) for l in sample_lines]
        if len(counts) > 0 and min(counts) >= 3 and max(counts) == min(counts):
            detected_delimiter = delim
            break
    
    if detected_delimiter is None:
        for delim in ['\t', ',', ';', ' ']:
            counts = [len(l.split(delim)) for l in sample_lines]
            if len(counts) > 0 and min(counts) >= 3:
                detected_delimiter = delim
                break
    
    if detected_delimiter is None:
        raise ValueError("Could not auto-detect file delimiter. Supported: tab, comma, semicolon, space")
    
    datetime_formats = ['%Y-%m-%d %H:%M:%S', '%d-%m-%Y %H:%M:%S',
                        '%m/%d/%Y %H:%M:%S', '%d/%m/%Y %H:%M:%S',
                        '%Y-%m-%d %H:%M', '%d-%m-%Y %H:%M']
    date_only_formats = ['%Y-%m-%d', '%d-%m-%Y', '%m-%d-%Y', '%d/%m/%Y', '%m/%d/%Y', 
                         '%Y/%m/%d', '%d.%m.%Y', '%Y.%m.%d']
    time_only_formats = ['%H:%M:%S', '%H:%M', '%I:%M:%S %p', '%I:%M %p']
    
    for line in raw_lines:
        line = line.strip()
        if not line:
            continue
        
        parts = line.split(detected_delimiter)
        parts = [p.strip() for p in parts if p.strip()]
        
        if len(parts) < 2:
            continue
        
        emp_id = None
        date_val = None
        time_val = None
        
        # Strategy 0: Combined datetime in one column
        for i, p in enumerate(parts):
            for fmt in datetime_formats:
                try:
                    dt = datetime.strptime(p, fmt)
                    date_val = dt.strftime('%Y-%m-%d')
                    time_val = dt.strftime('%H:%M:%S')
                    for j in range(len(parts)):
                        if j != i:
                            emp_id = parts[j]
                            break
                    break
                except ValueError:
                    continue
            if emp_id:
                break
        
        # Strategy 1: Separate date and time columns
        if not emp_id and len(parts) >= 3:
            try:
                emp_id_candidate = parts[0]
                for fmt in date_only_formats:
                    try:
                        dt = datetime.strptime(parts[1], fmt)
                        date_val = dt.strftime('%Y-%m-%d')
                        break
                    except ValueError:
                        continue
                if date_val:
                    for fmt in time_only_formats:
                        try:
                            datetime.strptime(parts[2], fmt)
                            time_val = parts[2]
                            break
                        except ValueError:
                            continue
                    if time_val:
                        emp_id = emp_id_candidate
                    else:
                        date_val = None
            except (IndexError, ValueError):
                pass
        
        # Strategy 2: Combine adjacent parts
        if not emp_id and len(parts) >= 3:
            for i in range(len(parts) - 1):
                combined = parts[i] + ' ' + parts[i+1]
                for fmt in datetime_formats:
                    try:
                        dt = datetime.strptime(combined, fmt)
                        date_val = dt.strftime('%Y-%m-%d')
                        time_val = dt.strftime('%H:%M:%S')
                        for j in range(len(parts)):
                            if j != i and j != i + 1:
                                emp_id = parts[j]
                                break
                        break
                    except ValueError:
                        continue
                if emp_id:
                    break
        
        if emp_id and date_val:
            rows.append({
                'emp_id': str(emp_id),
                'date': date_val,
                'time': time_val or '00:00:00'
            })
    
    if not rows:
        raise ValueError("Could not parse any valid attendance records from .dat file.\n"
                        "Expected format: emp_id, date, time (tab/comma/semicolon separated)")
    
    return pd.DataFrame(rows)


def load_attendance_file(filepath):
    """Load attendance from .xlsx, .xls, or .dat file. Returns DataFrame."""
    if filepath.lower().endswith('.dat'):
        return parse_dat_file(filepath)
    else:
        return pd.read_excel(filepath)


def compute_punch_times(att_df):
    """Compute first punch-in, last punch-out, and working hours per employee per date."""
    result = {}
    
    if 'time' not in att_df.columns:
        return result
    
    time_df = att_df.dropna(subset=['time'])
    if time_df.empty:
        return result
    
    time_df = time_df.copy()
    time_df['emp_id'] = time_df['emp_id'].astype(str)
    time_df['date_str'] = time_df['date'].dt.strftime('%Y-%m-%d') if pd.api.types.is_datetime64_any_dtype(time_df['date']) else time_df['date'].astype(str)
    
    for (eid, dstr), group in time_df.groupby(['emp_id', 'date_str']):
        times = []
        for t in group['time']:
            t_str = str(t).strip()
            if not t_str or t_str.lower() == 'nan':
                continue
            
            parsed_time = None
            for fmt in ['%H:%M:%S', '%H:%M', '%I:%M:%S %p', '%I:%M %p',
                        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
                        '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M',
                        '%m/%d/%Y %H:%M:%S', '%d/%m/%Y %H:%M:%S']:
                try:
                    parsed_time = datetime.strptime(t_str, fmt)
                    break
                except ValueError:
                    continue
            
            if parsed_time is None:
                try:
                    ts = pd.to_datetime(t_str)
                    parsed_time = ts.to_pydatetime()
                except:
                    pass
            
            if parsed_time is None:
                try:
                    if ':' in t_str:
                        time_part = t_str.split(' ')[-1] if ' ' in t_str else t_str
                        parts = time_part.split(':')
                        h = int(parts[0])
                        m = int(parts[1]) if len(parts) > 1 else 0
                        s = int(float(parts[2])) if len(parts) > 2 else 0
                        parsed_time = datetime(2000, 1, 1, h, m, s)
                except:
                    continue
            
            if parsed_time:
                times.append(parsed_time)
        
        if times:
            times.sort()
            punch_in = times[0]
            punch_out = times[-1]
            
            working_delta = punch_out - punch_in
            total_seconds = int(working_delta.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            
            result[(eid, dstr)] = {
                'punch_in': punch_in.strftime('%H:%M:%S'),
                'punch_out': punch_out.strftime('%H:%M:%S'),
                'working_hrs': f"{hours}h {minutes}m"
            }
    
    return result


# ================ EMPLOYEE MANAGEMENT FUNCTIONS ================

def refresh_employee_list():
    """Refresh the employee treeview"""
    global employee_df
    for item in emp_tree.get_children():
        emp_tree.delete(item)
    if employee_df is not None:
        for _, row in employee_df.iterrows():
            emp_tree.insert("", "end", values=(row['emp_id'], row['name'], row['Dept']))

def add_employee():
    """Add a new employee"""
    global employee_df
    
    dialog = tk.Toplevel(app)
    dialog.title("Add Employee")
    dialog.geometry("350x200")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()
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
        new_id = emp_id_entry.get().strip()
        new_name = name_entry.get().strip()
        new_dept = dept_entry.get().strip()
        
        if not new_id or not new_name or not new_dept:
            messagebox.showerror("Error", "All fields are required")
            return
        
        if employee_df is not None and new_id in employee_df['emp_id'].astype(str).values:
            messagebox.showerror("Error", "Employee ID already exists")
            return
        
        new_row = pd.DataFrame([{'emp_id': new_id, 'name': new_name, 'Dept': new_dept}])
        if employee_df is not None:
            employee_df = pd.concat([employee_df, new_row], ignore_index=True)
        else:
            employee_df = new_row
        
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
        
        if new_id != current_id and new_id in employee_df['emp_id'].astype(str).values:
            messagebox.showerror("Error", "Employee ID already exists")
            return
        
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

def upload_employee_list():
    """Upload/import employee list from Excel file"""
    global employee_df, emp_file
    
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return
    
    try:
        new_df = pd.read_excel(file_path)
        if not all(col in new_df.columns for col in ['emp_id', 'name', 'Dept']):
            messagebox.showerror("Error", "File must have columns: emp_id, name, Dept")
            return
        
        new_df = new_df[['emp_id', 'name', 'Dept']]
        new_df['emp_id'] = new_df['emp_id'].astype(str)
        
        if employee_df is not None:
            existing_ids = set(employee_df['emp_id'].astype(str).values)
            new_only = new_df[~new_df['emp_id'].astype(str).isin(existing_ids)]
            added = len(new_only)
            skipped = len(new_df) - added
            if added > 0:
                employee_df = pd.concat([employee_df, new_only], ignore_index=True)
            messagebox.showinfo("Upload Complete", f"Added {added} new employees\nSkipped {skipped} existing")
        else:
            employee_df = new_df
            emp_file = file_path
            messagebox.showinfo("Upload Complete", f"Loaded {len(new_df)} employees")
        
        refresh_employee_list()
        refresh_employee_dropdown()
    except Exception as e:
        messagebox.showerror("Error", f"Upload failed: {str(e)}")


# ================ UPLOAD ATTENDANCE FUNCTIONS ================

def browse_attendance_file():
    """Browse for attendance data file"""
    global att_file
    file_path = filedialog.askopenfilename(
        filetypes=[("All supported", "*.xlsx *.xls *.dat"), 
                   ("Excel files", "*.xlsx *.xls"), 
                   ("DAT files", "*.dat")]
    )
    if file_path:
        att_file = file_path
        upload_file_path_var.set(file_path)
        upload_status_label.config(text="File selected. Click 'Upload' to load.", fg="#333")

def upload_attendance_data():
    """Upload attendance data into memory"""
    global att_file, attendance_df
    
    if not att_file:
        messagebox.showerror("Error", "Please select an attendance file first")
        return
    
    try:
        attendance_df = load_attendance_file(att_file)
        attendance_df['emp_id'] = attendance_df['emp_id'].astype(str)
        
        record_count = len(attendance_df)
        upload_status_label.config(
            text=f"✅ Upload Successful! ({record_count} records loaded)", 
            fg="green", font=("Arial", 14, "bold")
        )
        upload_details_label.config(
            text=f"File: {os.path.basename(att_file)}\nRecords: {record_count}\nUploaded at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            fg="#555"
        )
    except Exception as e:
        upload_status_label.config(text="❌ Upload Failed!", fg="red", font=("Arial", 14, "bold"))
        upload_details_label.config(text=str(e), fg="red")
        messagebox.showerror("Error", f"Failed to load attendance file: {str(e)}")


# ================ GENERATE REPORT FUNCTIONS ================

def refresh_employee_dropdown():
    """Refresh employee and department dropdowns"""
    global employee_df
    if employee_df is not None:
        emp_list = [f"{row['emp_id']} - {row['name']}" for _, row in employee_df.iterrows()]
        emp_select_combo['values'] = emp_list
        editor_emp_combo['values'] = emp_list
        if emp_list:
            emp_select_combo.current(0)
            editor_emp_combo.current(0)
        depts = sorted(employee_df['Dept'].unique().tolist())
        dept_select_combo['values'] = depts
        if depts:
            dept_select_combo.current(0)
    else:
        emp_select_combo['values'] = []
        editor_emp_combo['values'] = []
        dept_select_combo['values'] = []

def on_report_type_change(*args):
    """Show/hide fields based on report type"""
    rtype = report_type_var.get()
    emp_select_frame.grid_remove()
    dept_select_frame.grid_remove()
    month_select_frame.grid_remove()
    custom_date_frame.grid_remove()
    
    if rtype == "Employee Wise":
        emp_select_frame.grid()
        custom_date_frame.grid()
    elif rtype == "Department Wise":
        dept_select_frame.grid()
        custom_date_frame.grid()
    elif rtype == "Month Wise":
        month_select_frame.grid()
    elif rtype == "Custom":
        custom_date_frame.grid()

def get_attendance_data():
    """Get attendance data from memory or file"""
    global attendance_df, att_file
    if attendance_df is not None:
        att = attendance_df.copy()
    elif att_file:
        att = load_attendance_file(att_file)
    else:
        return None
    att['emp_id'] = att['emp_id'].astype(str)
    att['date'] = pd.to_datetime(att['date'])
    return att

def build_attendance_base(emp, att, start, end):
    """Build base attendance DataFrame with status, applying saved edits"""
    all_dates = pd.date_range(start, end)
    punch_times = compute_punch_times(att)
    
    base = emp.copy()
    for d in all_dates:
        base[d.strftime("%Y-%m-%d")] = ""
    
    for d in all_dates:
        if d.weekday() in (5, 6):
            base[d.strftime("%Y-%m-%d")] = "WO"
        elif d.strftime("%Y-%m-%d") == "2025-12-25":
            base[d.strftime("%Y-%m-%d")] = "H"
    
    for _, r in att.iterrows():
        d_timestamp = r['date']
        if start <= d_timestamp <= end:
            d = d_timestamp.strftime("%Y-%m-%d")
            if d in base.columns:
                base.loc[base['emp_id'].astype(str) == str(r['emp_id']), d] = "P"
    
    for d in all_dates:
        col = d.strftime("%Y-%m-%d")
        base.loc[base[col] == "", col] = "A"
    
    # Apply saved edits (overrides everything)
    date_cols = [d.strftime("%Y-%m-%d") for d in all_dates]
    base = apply_edits_to_base(base, date_cols)
    
    return base, all_dates, punch_times

def write_excel_report(base, all_dates, punch_times, save_path, group_by_dept=True):
    """Write attendance report to Excel with styling and leave summary"""
    try:
        with pd.ExcelWriter(save_path, engine="openpyxl") as w:
            if group_by_dept:
                groups = base.groupby("Dept")
            else:
                groups = [("Report", base)]
            
            for dept, g in groups:
                g = g.copy()
                date_cols = [c for c in g.columns if c not in ["emp_id", "name", "Dept"]]
                
                # Summary columns
                g["Total_Present"] = (g[date_cols] == "P").sum(axis=1)
                g["Total_Absent"] = (g[date_cols] == "A").sum(axis=1)
                g["Total_WO"] = (g[date_cols] == "WO").sum(axis=1)
                g["Total_H"] = (g[date_cols] == "H").sum(axis=1)
                # Combined leave count
                g["Total_Leaves"] = sum((g[date_cols] == lt).sum(axis=1) for lt in LEAVE_TYPES)

                day_row_values = []
                for c in g.columns:
                    if c in date_cols:
                        day_row_values.append(datetime.strptime(c, "%Y-%m-%d").strftime("%A"))
                    else:
                        day_row_values.append("")

                punch_in_rows, punch_out_rows, working_hrs_rows = [], [], []
                for _, emp_row in g.iterrows():
                    pin, pout, wh = [], [], []
                    for c in g.columns:
                        if c in date_cols:
                            key = (str(emp_row['emp_id']), c)
                            if key in punch_times:
                                pin.append(punch_times[key]['punch_in'])
                                pout.append(punch_times[key]['punch_out'])
                                wh.append(punch_times[key]['working_hrs'])
                            else:
                                pin.append(""); pout.append(""); wh.append("")
                        elif c == 'name':
                            pin.append("Punch In"); pout.append("Punch Out"); wh.append("Working Hrs")
                        else:
                            pin.append(""); pout.append(""); wh.append("")
                    punch_in_rows.append(pin)
                    punch_out_rows.append(pout)
                    working_hrs_rows.append(wh)

                rows_list = [day_row_values]
                for idx, (_, emp_row) in enumerate(g.iterrows()):
                    rows_list.append(emp_row.tolist())
                    rows_list.append(punch_in_rows[idx])
                    rows_list.append(punch_out_rows[idx])
                    rows_list.append(working_hrs_rows[idx])

                final_df = pd.DataFrame(rows_list, columns=g.columns)
                sheet_name = str(dept)[:31]
                final_df.to_excel(w, sheet_name=sheet_name, index=False)

                ws = w.sheets[sheet_name]
                red_font = Font(color="FF0000")
                blue_font = Font(color="0000FF", italic=True, size=9)
                green_font = Font(color="006600", italic=True, size=9)
                gray_font = Font(color="666666", italic=True, size=9)

                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value == "WO":
                            cell.font = red_font

                num_emps = len(g)
                for i in range(num_emps):
                    punch_in_row = 3 + (i * 4) + 1
                    punch_out_row = punch_in_row + 1
                    working_hrs_row = punch_out_row + 1
                    for row_num, font_style in [(punch_in_row, blue_font),
                                                 (punch_out_row, green_font),
                                                 (working_hrs_row, gray_font)]:
                        try:
                            for cell in ws[row_num]:
                                if cell.value and cell.value not in ('', 'Punch In', 'Punch Out', 'Working Hrs'):
                                    cell.font = font_style
                                elif cell.value in ('Punch In', 'Punch Out', 'Working Hrs'):
                                    cell.font = Font(bold=True, size=9)
                        except:
                            pass

        messagebox.showinfo("Success", f"Report generated successfully!\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

def generate_report():
    """Generate report based on selected type — Excel export only"""
    global employee_df
    
    if employee_df is None:
        messagebox.showerror("Error", "Please load employee data first (from Employee tab)")
        return
    
    att = get_attendance_data()
    if att is None:
        messagebox.showerror("Error", "Please upload attendance data first (from Upload tab)")
        return
    
    rtype = report_type_var.get()
    emp = employee_df.copy()
    emp = emp[['emp_id', 'name', 'Dept']]
    emp['emp_id'] = emp['emp_id'].astype(str)
    
    try:
        if rtype == "Employee Wise":
            selected_emp = emp_select_var.get()
            if not selected_emp:
                messagebox.showerror("Error", "Please select an employee")
                return
            emp_id = selected_emp.split(" - ")[0].strip()
            emp = emp[emp['emp_id'] == emp_id]
            
            start_str = gen_start_date_var.get()
            end_str = gen_end_date_var.get()
            if start_str and end_str:
                start = pd.to_datetime(start_str)
                end = pd.to_datetime(end_str)
            else:
                emp_att = att[att['emp_id'] == emp_id]
                if len(emp_att) == 0:
                    messagebox.showinfo("Info", "No attendance records for this employee")
                    return
                start = emp_att['date'].min()
                end = emp_att['date'].max()
            
            base, all_dates, punch_times = build_attendance_base(emp, att, start, end)
            save = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Employee Report")
            if save:
                write_excel_report(base, all_dates, punch_times, save, group_by_dept=False)
        
        elif rtype == "Department Wise":
            selected_dept = dept_select_var.get()
            if not selected_dept:
                messagebox.showerror("Error", "Please select a department")
                return
            emp = emp[emp['Dept'] == selected_dept]
            
            start_str = gen_start_date_var.get()
            end_str = gen_end_date_var.get()
            if start_str and end_str:
                start = pd.to_datetime(start_str)
                end = pd.to_datetime(end_str)
            else:
                dept_att = att[att['emp_id'].isin(emp['emp_id'])]
                if len(dept_att) == 0:
                    messagebox.showinfo("Info", "No attendance records for this department")
                    return
                start = dept_att['date'].min()
                end = dept_att['date'].max()
            
            base, all_dates, punch_times = build_attendance_base(emp, att, start, end)
            save = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Department Report")
            if save:
                write_excel_report(base, all_dates, punch_times, save, group_by_dept=False)
        
        elif rtype == "Month Wise":
            try:
                year = int(month_year_var.get())
                month_str = month_month_var.get()
                month = int(month_str.split(' - ')[0])
            except (ValueError, IndexError):
                messagebox.showerror("Error", "Please select valid year and month")
                return
            
            start = pd.Timestamp(year, month, 1)
            end = start + pd.offsets.MonthEnd(1)
            
            base, all_dates, punch_times = build_attendance_base(emp, att, start, end)
            save = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 title=f"Save Monthly Report - {year}-{month:02d}")
            if save:
                write_excel_report(base, all_dates, punch_times, save, group_by_dept=True)
        
        elif rtype == "Custom":
            start_str = gen_start_date_var.get()
            end_str = gen_end_date_var.get()
            if not start_str or not end_str:
                messagebox.showerror("Error", "Please select both From and To dates")
                return
            start = pd.to_datetime(start_str)
            end = pd.to_datetime(end_str)
            
            base, all_dates, punch_times = build_attendance_base(emp, att, start, end)
            save = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Custom Report")
            if save:
                write_excel_report(base, all_dates, punch_times, save, group_by_dept=True)
    
    except Exception as e:
        messagebox.showerror("Error", f"Report generation failed: {str(e)}")


def generate_print_report():
    """Generate a printer-friendly PDF report in A4 landscape mode"""
    global employee_df

    if employee_df is None:
        messagebox.showerror("Error", "Please load employee data first (from Employee tab)")
        return

    att = get_attendance_data()
    if att is None:
        messagebox.showerror("Error", "Please upload attendance data first (from Upload tab)")
        return

    rtype = report_type_var.get()
    emp = employee_df.copy()
    emp = emp[['emp_id', 'name', 'Dept']]
    emp['emp_id'] = emp['emp_id'].astype(str)

    try:
        if rtype == "Employee Wise":
            selected_emp = emp_select_var.get()
            if not selected_emp:
                messagebox.showerror("Error", "Please select an employee")
                return
            emp_id = selected_emp.split(" - ")[0].strip()
            emp = emp[emp['emp_id'] == emp_id]
            start_str = gen_start_date_var.get()
            end_str = gen_end_date_var.get()
            if start_str and end_str:
                start = pd.to_datetime(start_str)
                end = pd.to_datetime(end_str)
            else:
                emp_att = att[att['emp_id'] == emp_id]
                if len(emp_att) == 0:
                    messagebox.showinfo("Info", "No attendance records for this employee")
                    return
                start = emp_att['date'].min()
                end = emp_att['date'].max()
            title = f"Attendance Report - {selected_emp}"

        elif rtype == "Department Wise":
            selected_dept = dept_select_var.get()
            if not selected_dept:
                messagebox.showerror("Error", "Please select a department")
                return
            emp = emp[emp['Dept'] == selected_dept]
            start_str = gen_start_date_var.get()
            end_str = gen_end_date_var.get()
            if start_str and end_str:
                start = pd.to_datetime(start_str)
                end = pd.to_datetime(end_str)
            else:
                dept_att = att[att['emp_id'].isin(emp['emp_id'])]
                if len(dept_att) == 0:
                    messagebox.showinfo("Info", "No attendance records for this department")
                    return
                start = dept_att['date'].min()
                end = dept_att['date'].max()
            title = f"Attendance Report - Dept: {selected_dept}"

        elif rtype == "Month Wise":
            try:
                year = int(month_year_var.get())
                month_str = month_month_var.get()
                month = int(month_str.split(' - ')[0])
            except (ValueError, IndexError):
                messagebox.showerror("Error", "Please select valid year and month")
                return
            start = pd.Timestamp(year, month, 1)
            end = start + pd.offsets.MonthEnd(1)
            title = f"Attendance Report - {datetime(year, month, 1).strftime('%B %Y')}"

        elif rtype == "Custom":
            start_str = gen_start_date_var.get()
            end_str = gen_end_date_var.get()
            if not start_str or not end_str:
                messagebox.showerror("Error", "Please select both From and To dates")
                return
            start = pd.to_datetime(start_str)
            end = pd.to_datetime(end_str)
            title = "Attendance Report - Custom Range"
        else:
            return

        base, all_dates, punch_times = build_attendance_base(emp, att, start, end)
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Print Report"
        )
        if not save:
            return

        write_pdf_report(base, all_dates, punch_times, save, title, start, end)

    except Exception as e:
        messagebox.showerror("Error", f"Print report generation failed: {str(e)}")


def write_pdf_report(base, all_dates, punch_times, save_path, title, start, end):
    """Write a printer-friendly PDF report in A4 landscape mode"""
    try:
        page_w, page_h = landscape(A4)
        doc = SimpleDocTemplate(
            save_path,
            pagesize=landscape(A4),
            leftMargin=10*mm,
            rightMargin=10*mm,
            topMargin=12*mm,
            bottomMargin=12*mm
        )

        elements = []
        styles = getSampleStyleSheet()

        # Title
        title_style = ParagraphStyle(
            'ReportTitle', parent=styles['Title'],
            fontSize=14, alignment=TA_CENTER, spaceAfter=4
        )
        elements.append(Paragraph(title, title_style))

        # Date range subtitle
        subtitle = f"Period: {start.strftime('%d-%b-%Y')} to {end.strftime('%d-%b-%Y')}  |  Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}"
        sub_style = ParagraphStyle(
            'Subtitle', parent=styles['Normal'],
            fontSize=8, alignment=TA_CENTER, spaceAfter=8, textColor=colors.black
        )
        elements.append(Paragraph(subtitle, sub_style))

        date_cols = [d.strftime("%Y-%m-%d") for d in all_dates]
        num_dates = len(date_cols)

        # Compute summary counts per employee
        summary_data = {}
        for _, row in base.iterrows():
            eid = str(row['emp_id'])
            p_count = sum(1 for c in date_cols if str(row.get(c, '')) == 'P')
            a_count = sum(1 for c in date_cols if str(row.get(c, '')) == 'A')
            wo_count = sum(1 for c in date_cols if str(row.get(c, '')) == 'WO')
            h_count = sum(1 for c in date_cols if str(row.get(c, '')) == 'H')
            leave_count = sum(1 for c in date_cols if str(row.get(c, '')) in LEAVE_TYPES)
            summary_data[eid] = (p_count, a_count, wo_count, h_count, leave_count)

        # Build table header
        header_row1 = ['ID', 'Name', 'Dept']
        header_row2 = ['', '', '']
        for d in all_dates:
            header_row1.append(d.strftime('%d'))
            header_row2.append(d.strftime('%a')[:2])
        header_row1 += ['P', 'A', 'WO', 'H', 'LV']
        header_row2 += ['', '', '', '', '']

        # Build data rows
        data_rows = [header_row1, header_row2]
        for _, row in base.iterrows():
            eid = str(row['emp_id'])
            data_row = [eid, str(row['name'])[:18], str(row['Dept'])[:10]]
            for c in date_cols:
                data_row.append(str(row.get(c, '')))
            s = summary_data.get(eid, (0, 0, 0, 0, 0))
            data_row += [str(s[0]), str(s[1]), str(s[2]), str(s[3]), str(s[4])]
            data_rows.append(data_row)

        # Calculate column widths to fit A4 landscape
        usable_w = page_w - 20*mm
        id_w = 32
        name_w = 62
        dept_w = 45
        summary_w = 18
        fixed_w = id_w + name_w + dept_w + (5 * summary_w)
        remaining_w = usable_w - fixed_w
        date_col_w = max(14, remaining_w / max(num_dates, 1))

        col_widths = [id_w, name_w, dept_w]
        col_widths += [date_col_w] * num_dates
        col_widths += [summary_w] * 5

        # Auto-scale font
        if num_dates <= 15:
            font_size = 7
        elif num_dates <= 25:
            font_size = 6
        else:
            font_size = 5

        table = Table(data_rows, colWidths=col_widths, repeatRows=2)

        # Black and white table style
        style_commands = [
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), font_size),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 1), font_size),
            ('BACKGROUND', (0, 0), (-1, 1), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 1), colors.white),
            ('TEXTCOLOR', (0, 2), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 2), (1, -1), 'LEFT'),
            ('ALIGN', (2, 2), (2, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('LINEBELOW', (0, 1), (-1, 1), 1.2, colors.black),
            ('ROWBACKGROUNDS', (0, 2), (-1, -1), [colors.white, colors.Color(0.93, 0.93, 0.93)]),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            # Vertical line before summary
            ('LINEAFTER', (2, 0), (2, -1), 1, colors.black),
            ('LINEBEFORE', (3 + num_dates, 0), (3 + num_dates, -1), 1, colors.black),
            # Bold summary columns
            ('FONTNAME', (3 + num_dates, 2), (-1, -1), 'Helvetica-Bold'),
        ]

        table.setStyle(TableStyle(style_commands))
        elements.append(table)

        # Legend
        elements.append(Spacer(1, 6*mm))
        legend_style = ParagraphStyle(
            'Legend', parent=styles['Normal'],
            fontSize=7, textColor=colors.black, alignment=TA_LEFT
        )
        legend_text = ("<b>Legend:</b>  P = Present  |  A = Absent  |  WO = Weekly Off  |  "
                       "H = Holiday  |  LV = Total Leaves (EL, CL, COFF, DL, ML, HD, CC, MaL, RH)")
        elements.append(Paragraph(legend_text, legend_style))

        # Build with page numbers
        def add_page_number(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 7)
            canvas.setFillColor(colors.black)
            canvas.drawRightString(
                page_w - 10*mm, 7*mm,
                f"Page {doc.page}"
            )
            canvas.drawString(
                10*mm, 7*mm,
                "IOP Attendance Report"
            )
            canvas.restoreState()

        doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)
        messagebox.showinfo("Success", f"Print report generated successfully!\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate print report: {str(e)}")


# ================ EDITOR FUNCTIONS ================

def verify_pin():
    """Ask for PIN before allowing edit access"""
    pin = simpledialog.askstring("Security", "Enter PIN to edit:", show='*', parent=app)
    if pin is None:
        return False
    if pin == EDITOR_PIN:
        return True
    else:
        messagebox.showerror("Access Denied", "Incorrect PIN!")
        return False

def view_employee_attendance():
    """View attendance for selected employee in editor (with saved edits applied)"""
    global attendance_df, att_file, employee_df
    
    att = get_attendance_data()
    if att is None:
        messagebox.showerror("Error", "Please upload attendance data first (from Upload tab)")
        return
    
    selected_emp = editor_emp_var.get()
    if not selected_emp:
        messagebox.showerror("Error", "Please select an employee")
        return
    
    emp_id = selected_emp.split(" - ")[0].strip()
    start_str = editor_start_var.get()
    end_str = editor_end_var.get()
    
    try:
        punch_times = compute_punch_times(att)
        emp_att = att[att['emp_id'].astype(str) == str(emp_id)]
        saved_edits = load_saved_edits()
        
        if start_str and end_str:
            start = pd.to_datetime(start_str)
            end = pd.to_datetime(end_str)
        else:
            if len(emp_att) > 0:
                start = emp_att['date'].min()
                end = emp_att['date'].max()
            else:
                messagebox.showinfo("Info", "No attendance records found for this employee")
                return
        
        all_dates = pd.date_range(start, end)
        
        for item in att_tree.get_children():
            att_tree.delete(item)
        
        present_dates = set(emp_att['date'].dt.strftime("%Y-%m-%d").values)
        counts = {}
        
        for d in all_dates:
            d_str = d.strftime("%Y-%m-%d")
            day_name = d.strftime("%A")
            
            # Check if there's a saved edit for this emp+date
            edit_key = f"{emp_id}_{d_str}"
            saved = saved_edits.get(edit_key, None)
            
            if saved:
                status = saved
            elif d.weekday() in (5, 6):
                status = "WO"
            elif d_str == "2025-12-25":
                status = "H"
            elif d_str in present_dates:
                status = "P"
            else:
                status = "A"
            
            counts[status] = counts.get(status, 0) + 1
            
            key = (str(emp_id), d_str)
            punch_in = punch_times.get(key, {}).get('punch_in', '')
            punch_out = punch_times.get(key, {}).get('punch_out', '')
            working_hrs = punch_times.get(key, {}).get('working_hrs', '')
            
            att_tree.insert("", "end", values=(d_str, day_name, status, punch_in, punch_out, working_hrs))
        
        # Build summary string
        summary_parts = []
        for s in ['P', 'A', 'WO', 'H'] + LEAVE_TYPES:
            if counts.get(s, 0) > 0:
                summary_parts.append(f"{s}: {counts[s]}")
        editor_summary_label.config(text=" | ".join(summary_parts))
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load attendance: {str(e)}")

def edit_attendance_status():
    """Edit attendance status — PIN protected"""
    if not verify_pin():
        return
    
    selected = att_tree.selection()
    if not selected:
        messagebox.showerror("Error", "Please select a row to edit")
        return
    
    item = att_tree.item(selected[0])
    values = list(item['values'])
    current_date = values[0]
    current_status = values[2]
    
    dialog = tk.Toplevel(app)
    dialog.title(f"Edit Attendance - {current_date}")
    dialog.geometry("350x200")
    dialog.resizable(False, False)
    dialog.transient(app)
    dialog.grab_set()
    dialog.geometry("+%d+%d" % (app.winfo_x() + 150, app.winfo_y() + 150))
    
    tk.Label(dialog, text=f"Date: {current_date}", font=("Arial", 11, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
    tk.Label(dialog, text=f"Current Status: {current_status}", font=("Arial", 10)).grid(row=1, column=0, columnspan=2, pady=5)
    
    tk.Label(dialog, text="New Status:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    status_var = tk.StringVar(value=current_status)
    status_combo = ttk.Combobox(dialog, textvariable=status_var, values=ALL_STATUS_OPTIONS, width=15, state="readonly")
    status_combo.grid(row=2, column=1, padx=10, pady=10)
    
    def save_edit():
        new_status = status_var.get()
        values[2] = new_status
        att_tree.item(selected[0], values=values)
        dialog.destroy()
    
    btn_frame = tk.Frame(dialog)
    btn_frame.grid(row=3, column=0, columnspan=2, pady=15)
    tk.Button(btn_frame, text="Save", command=save_edit, bg="green", fg="white", width=10).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=10).pack(side="left", padx=5)

def save_all_edits():
    """Save all changes from treeview to persistent JSON file"""
    items = att_tree.get_children()
    if not items:
        messagebox.showerror("Error", "No attendance data to save")
        return
    
    selected_emp = editor_emp_var.get()
    if not selected_emp:
        messagebox.showerror("Error", "No employee selected")
        return
    
    emp_id = selected_emp.split(" - ")[0].strip()
    
    # Load existing edits
    edits = load_saved_edits()
    
    # Save current treeview data
    save_count = 0
    for item in items:
        values = att_tree.item(item)['values']
        date_str = str(values[0])
        status = str(values[2])
        key = f"{emp_id}_{date_str}"
        
        # Only save if not already saved (can't override existing edits)
        if key not in edits:
            edits[key] = status
            save_count += 1
        elif edits[key] != status:
            # Allow updating if content changed in this session
            edits[key] = status
            save_count += 1
    
    if save_edits_to_file(edits):
        messagebox.showinfo("Success", f"Attendance edits saved!\n{save_count} records saved to file.\nThese will be reflected in Generate Report.")


# ================ GUI SETUP ================

app = tk.Tk()
app.title("IOP Attendance Generator")
app.geometry("900x700")

notebook = ttk.Notebook(app)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# ================ TAB 1: EMPLOYEE ================
tab_employees = ttk.Frame(notebook)
notebook.add(tab_employees, text="Employee")

tree_frame = tk.Frame(tab_employees)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

emp_tree = ttk.Treeview(tree_frame, columns=("ID", "Name", "Department"), show="headings", height=15)
emp_tree.heading("ID", text="Employee ID")
emp_tree.heading("Name", text="Name")
emp_tree.heading("Department", text="Department")
emp_tree.column("ID", width=120)
emp_tree.column("Name", width=250)
emp_tree.column("Department", width=200)

scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=emp_tree.yview)
emp_tree.configure(yscrollcommand=scrollbar.set)
emp_tree.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

btn_frame = tk.Frame(tab_employees)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Add Employee", command=add_employee, bg="#4CAF50", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Edit Employee", command=edit_employee, bg="#2196F3", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Delete Employee", command=delete_employee, bg="#f44336", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Upload List", command=upload_employee_list, bg="#9C27B0", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(btn_frame, text="Save to File", command=save_employee_file, bg="#FF9800", fg="white", 
          width=15, font=("Arial", 10)).pack(side="left", padx=5)

# ================ TAB 2: UPLOAD ATTENDANCE ================
tab_upload = ttk.Frame(notebook)
notebook.add(tab_upload, text="Upload Attendance")

upload_frame = tk.Frame(tab_upload)
upload_frame.pack(pady=40)

tk.Label(upload_frame, text="Upload Attendance Data", font=("Arial", 16, "bold"), fg="#333").pack(pady=15)
tk.Label(upload_frame, text="Select attendance file (.xlsx, .xls, or .dat) and click Upload", 
         font=("Arial", 10), fg="#666").pack(pady=5)

file_select_frame = tk.Frame(upload_frame)
file_select_frame.pack(pady=20)

upload_file_path_var = tk.StringVar()
tk.Entry(file_select_frame, textvariable=upload_file_path_var, width=50, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(file_select_frame, text="Browse", command=browse_attendance_file, bg="#607D8B", fg="white",
          font=("Arial", 10), width=10).pack(side="left", padx=5)

tk.Button(upload_frame, text="📤  Upload", command=upload_attendance_data, bg="#4CAF50", fg="white", 
          font=("Arial", 14, "bold"), padx=30, pady=8).pack(pady=20)

upload_status_label = tk.Label(upload_frame, text="", font=("Arial", 12), fg="#333")
upload_status_label.pack(pady=10)

upload_details_label = tk.Label(upload_frame, text="", font=("Arial", 10), fg="#555", justify="center")
upload_details_label.pack(pady=5)


# ================ TAB 3: EDITOR ================
tab_editor = ttk.Frame(notebook)
notebook.add(tab_editor, text="Editor")

editor_select_frame = tk.Frame(tab_editor)
editor_select_frame.pack(pady=10, padx=10, fill="x")

tk.Label(editor_select_frame, text="Employee:", font=("Arial", 10)).grid(row=0, column=0, padx=5, pady=5, sticky="e")
editor_emp_var = tk.StringVar()
editor_emp_combo = ttk.Combobox(editor_select_frame, textvariable=editor_emp_var, width=35, state="readonly")
editor_emp_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

editor_start_var = tk.StringVar()
editor_end_var = tk.StringVar()

tk.Label(editor_select_frame, text="From:", font=("Arial", 10)).grid(row=0, column=2, padx=5, pady=5, sticky="e")
editor_start_entry = DateEntry(editor_select_frame, textvariable=editor_start_var, width=12,
                                date_pattern='yyyy-mm-dd', font=('Arial', 10))
editor_start_entry.delete(0, 'end')
editor_start_entry.grid(row=0, column=3, padx=5, pady=5)

tk.Label(editor_select_frame, text="To:", font=("Arial", 10)).grid(row=0, column=4, padx=5, pady=5, sticky="e")
editor_end_entry = DateEntry(editor_select_frame, textvariable=editor_end_var, width=12,
                              date_pattern='yyyy-mm-dd', font=('Arial', 10))
editor_end_entry.delete(0, 'end')
editor_end_entry.grid(row=0, column=5, padx=5, pady=5)

tk.Button(editor_select_frame, text="View", command=view_employee_attendance, bg="#673AB7", fg="white",
          font=("Arial", 10, "bold"), width=8).grid(row=0, column=6, padx=10, pady=5)

# Attendance Treeview
att_tree_frame = tk.Frame(tab_editor)
att_tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

att_tree = ttk.Treeview(att_tree_frame,
                        columns=("Date", "Day", "Status", "PunchIn", "PunchOut", "WorkingHrs"),
                        show="headings", height=14)
att_tree.heading("Date", text="Date")
att_tree.heading("Day", text="Day")
att_tree.heading("Status", text="Status")
att_tree.heading("PunchIn", text="Punch In")
att_tree.heading("PunchOut", text="Punch Out")
att_tree.heading("WorkingHrs", text="Working Hrs")
att_tree.column("Date", width=110)
att_tree.column("Day", width=100)
att_tree.column("Status", width=70)
att_tree.column("PunchIn", width=90)
att_tree.column("PunchOut", width=90)
att_tree.column("WorkingHrs", width=90)

att_scrollbar = ttk.Scrollbar(att_tree_frame, orient="vertical", command=att_tree.yview)
att_tree.configure(yscrollcommand=att_scrollbar.set)
att_tree.pack(side="left", fill="both", expand=True)
att_scrollbar.pack(side="right", fill="y")

editor_summary_label = tk.Label(tab_editor, text="", font=("Arial", 11, "bold"), fg="#333")
editor_summary_label.pack(pady=3)

editor_btn_frame = tk.Frame(tab_editor)
editor_btn_frame.pack(pady=5)

tk.Button(editor_btn_frame, text="🔒 Edit Status", command=edit_attendance_status, bg="#2196F3", fg="white",
          width=15, font=("Arial", 10)).pack(side="left", padx=5)
tk.Button(editor_btn_frame, text="💾 Save Changes", command=save_all_edits, bg="#4CAF50", fg="white",
          width=15, font=("Arial", 10)).pack(side="left", padx=5)

att_tree.bind("<Double-1>", lambda e: edit_attendance_status())


# ================ TAB 4: GENERATE REPORT ================
tab_generate = ttk.Frame(notebook)
notebook.add(tab_generate, text="Generate Report")

type_frame = tk.LabelFrame(tab_generate, text="Report Type", font=("Arial", 10, "bold"))
type_frame.pack(fill="x", padx=10, pady=10)

report_type_var = tk.StringVar(value="Employee Wise")
for rtype in ["Employee Wise", "Department Wise", "Month Wise", "Custom"]:
    tk.Radiobutton(type_frame, text=rtype, variable=report_type_var, value=rtype,
                   font=("Arial", 10)).pack(side="left", padx=15, pady=8)

options_frame = tk.Frame(tab_generate)
options_frame.pack(fill="x", padx=10, pady=5)

emp_select_frame = tk.Frame(options_frame)
emp_select_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky="w")
tk.Label(emp_select_frame, text="Select Employee:", font=("Arial", 10)).pack(side="left", padx=5)
emp_select_var = tk.StringVar()
emp_select_combo = ttk.Combobox(emp_select_frame, textvariable=emp_select_var, width=40, state="readonly")
emp_select_combo.pack(side="left", padx=5)

dept_select_frame = tk.Frame(options_frame)
dept_select_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky="w")
tk.Label(dept_select_frame, text="Select Department:", font=("Arial", 10)).pack(side="left", padx=5)
dept_select_var = tk.StringVar()
dept_select_combo = ttk.Combobox(dept_select_frame, textvariable=dept_select_var, width=40, state="readonly")
dept_select_combo.pack(side="left", padx=5)

month_select_frame = tk.Frame(options_frame)
month_select_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky="w")

tk.Label(month_select_frame, text="Year:", font=("Arial", 10)).pack(side="left", padx=5)
current_year = datetime.now().year
year_options = [str(y) for y in range(current_year - 5, current_year + 2)]
month_year_var = tk.StringVar(value=str(current_year))
ttk.Combobox(month_select_frame, textvariable=month_year_var, values=year_options, 
             width=8, state="readonly").pack(side="left", padx=5)

tk.Label(month_select_frame, text="Month:", font=("Arial", 10)).pack(side="left", padx=5)
month_options = [f"{m} - {datetime(2000, m, 1).strftime('%B')}" for m in range(1, 13)]
month_month_var = tk.StringVar(value=f"{datetime.now().month} - {datetime.now().strftime('%B')}")
ttk.Combobox(month_select_frame, textvariable=month_month_var, values=month_options, 
             width=15, state="readonly").pack(side="left", padx=5)

custom_date_frame = tk.Frame(options_frame)
custom_date_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky="w")

gen_start_date_var = tk.StringVar()
gen_end_date_var = tk.StringVar()

tk.Label(custom_date_frame, text="From:", font=("Arial", 10)).pack(side="left", padx=5)
gen_start_date_entry = DateEntry(custom_date_frame, textvariable=gen_start_date_var, width=15,
                                  date_pattern='yyyy-mm-dd', font=('Arial', 10))
gen_start_date_entry.delete(0, 'end')
gen_start_date_entry.pack(side="left", padx=5)

tk.Label(custom_date_frame, text="To:", font=("Arial", 10)).pack(side="left", padx=5)
gen_end_date_entry = DateEntry(custom_date_frame, textvariable=gen_end_date_var, width=15,
                                date_pattern='yyyy-mm-dd', font=('Arial', 10))
gen_end_date_entry.delete(0, 'end')
gen_end_date_entry.pack(side="left", padx=5)

gen_btn_frame = tk.Frame(tab_generate)
gen_btn_frame.pack(pady=30)

tk.Button(gen_btn_frame, text="📊  Generate Report", command=generate_report, bg="#673AB7", fg="white",
          font=("Arial", 14, "bold"), padx=30, pady=10).pack(side="left", padx=10)
tk.Button(gen_btn_frame, text="🖨  Print Report", command=generate_print_report, bg="#00796B", fg="white",
          font=("Arial", 14, "bold"), padx=30, pady=10).pack(side="left", padx=10)

tk.Label(tab_generate, text="Select report type, fill in options, and click Generate.\nEdited attendance data will be reflected in reports.", 
         font=("Arial", 10), fg="#666", justify="center").pack(pady=10)

report_type_var.trace_add("write", on_report_type_change)
on_report_type_change()

# Footer
tk.Label(app, text="Developed & Designed @ HEPD Lab by SIDHARTHA SAHU", font=("Times New Roman", 12), fg="#999").pack(side="bottom", pady=2)

app.mainloop()
