# Quick Start Guide - Attendance Management System

## ⚠️ No Data Showing? Follow These Steps:

### Step 1: Add Employees First (Required!)

**Option A: Add Manually**
1. Go to **"👥 Employees"** tab
2. Click **"➕ Add Employee"**
3. Fill in: Employee ID, Name, Department
4. Click Save
5. Repeat for all employees

**Option B: Import from Excel (Faster for multiple employees)**
1. Create Excel file with 3 columns:
   - `emp_id` (e.g., E001, E002)
   - `name` (e.g., John Doe)
   - `Dept` (e.g., IT, HR)
2. Go to **"👥 Employees"** tab
3. Click **"📤 Import from Excel"**
4. Select your file
5. All employees added at once!

**Sample Employees Excel:**
```
emp_id  | name        | Dept
--------|-------------|-------
E001    | John Doe    | IT
E002    | Jane Smith  | HR
E003    | Bob Johnson | IT
E004    | Alice Brown | Finance
E005    | Charlie W   | HR
```

---

### Step 2: Upload Attendance Data

**Required: 3-Column Format**
1. Create Excel/CSV with exactly 3 columns:
   - `emp_id` - Employee ID (must match employees added in Step 1)
   - `date` - Date in YYYY-MM-DD format
   - `time` - Time in HH:MM:SS format

2. Go to **"📋 Attendance"** tab
3. Click **"📁 Choose File & Upload"**
4. Select your file
5. System will show upload summary

**Sample Attendance Excel:**
```
emp_id  | date       | time
--------|------------|----------
E001    | 2026-02-01 | 09:00:00
E001    | 2026-02-01 | 18:00:00
E002    | 2026-02-01 | 09:15:00
E002    | 2026-02-01 | 17:45:00
```

---

### Step 3: View Attendance Data

**Option 1: Generate Report**
1. Go to **"📊 Reports"** tab
2. Select report type:
   - Employee Report (single employee)
   - Department Report (all in dept)
   - Monthly Report (full month)
3. Click **"📊 Generate Report"**
4. Save Excel file
5. Open to view data

**Option 2: Export All Data**
1. Go to **"⚙️ Settings"** tab
2. Click **"📤 Export All to Excel"**
3. Opens multi-sheet Excel with all data

---

## 🎯 Quick Test with Sample Data

**Use the provided sample files:**

1. **Add 5 sample employees:**
   - Run: `python create_sample_attendance.py`
   - This creates `sample_3col_attendance.xlsx`
   - Contains 5 employees (E001-E005)

2. **Create employees first:**
   - Manually add E001-E005 in Employees tab, OR
   - Import from Excel if you have employee file

3. **Upload attendance:**
   - Use Attendance tab
   - Upload `sample_3col_attendance.xlsx`
   - Should show: "✅ New records added: 50"

4. **View data:**
   - Go to Reports tab
   - Select "Monthly Report"
   - Year: 2026, Month: 2
   - Generate and save

---

## ❗ Common Issues

**Issue: "No data showing in reports"**
- ✅ Fix: Make sure you added employees first!
- ✅ Fix: Upload attendance data second
- ✅ Fix: Check Reports tab, not Attendance tab

**Issue: "Upload failed - missing employee"**
- ✅ Fix: Employee IDs in attendance file must match employee IDs in system
- ✅ Fix: Add employees before uploading attendance

**Issue: "Duplicates skipped"**
- ✅ This is normal! System prevents duplicate entries
- ✅ Only new data is added

---

## 📊 Data Flow

```
1. Add Employees → Database
2. Upload Attendance → Database (matches employee IDs)
3. Generate Reports → Pulls from database
```

**Important:** Attendance records MUST have matching employee IDs!

---

## ✅ Checklist

- [ ] Step 1: Add employees (Employees tab)
- [ ] Step 2: Upload attendance (Attendance tab)  
- [ ] Step 3: Generate report (Reports tab)
- [ ] Step 4: View exported Excel file

Follow this order!
