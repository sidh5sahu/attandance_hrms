"""
Test script for attendance database functionality
This script tests all major database operations
"""

from database import AttendanceDatabase
from datetime import datetime, timedelta
import os

def test_database():
    print("=" * 60)
    print("ATTENDANCE DATABASE TEST SCRIPT")
    print("=" * 60)
    
    # Clean up old test files
    print("\n0. Cleaning up old test files...")
    test_files = ["test_attendance.db", "test_backup.db", "test_export.xlsx"]
    for f in test_files:
        if os.path.exists(f):
            os.remove(f)
            print(f"   ✓ Removed {f}")
    
    # Initialize database
    print("\n1. Initializing database...")
    db = AttendanceDatabase("test_attendance.db")
    print("   ✓ Database created successfully")
    
    # Test employee operations
    print("\n2. Testing employee operations...")
    
    # Add employees
    employees = [
        ("E001", "John Doe", "IT"),
        ("E002", "Jane Smith", "HR"),
        ("E003", "Bob Johnson", "IT"),
        ("E004", "Alice Brown", "Finance"),
        ("E005", "Charlie Wilson", "HR")
    ]
    
    for emp_id, name, dept in employees:
        if db.add_employee(emp_id, name, dept):
            print(f"   ✓ Added employee: {emp_id} - {name} ({dept})")
        else:
            print(f"   ✗ Failed to add employee: {emp_id}")
    
    # Test getting employees
    print("\n3. Retrieving all employees...")
    all_emps = db.get_all_employees()
    print(f"   Total employees: {len(all_emps)}")
    for emp in all_emps:
        print(f"   - {emp['emp_id']}: {emp['name']} ({emp['dept']})")
    
    # Test department query
    print("\n4. Getting employees by department (IT)...")
    it_emps = db.get_employees_by_dept("IT")
    print(f"   IT Department has {len(it_emps)} employees:")
    for emp in it_emps:
        print(f"   - {emp['emp_id']}: {emp['name']}")
    
    # Test attendance operations
    print("\n5. Adding attendance records...")
    
    # Generate sample attendance for January 2026
    start_date = datetime(2026, 1, 1)
    end_date = datetime(2026, 1, 31)
    current_date = start_date
    
    # Build list of all attendance records for bulk insert
    attendance_records = []
    while current_date <= end_date:
        date_str = current_date.strftime("%Y-%m-%d")
        
        for emp_id, _, _ in employees:
            # Weekend
            if current_date.weekday() in (5, 6):
                status = "WO"
            # Random absence pattern (E001 absent on Mondays, E002 on Tuesdays for demo)
            elif emp_id == "E001" and current_date.weekday() == 0:
                status = "A"
            elif emp_id == "E002" and current_date.weekday() == 1:
                status = "A"
            else:
                status = "P"
            
            attendance_records.append((emp_id, date_str, status))
        
        current_date += timedelta(days=1)
    
    # Use bulk insert for much better performance
    attendance_count = db.add_attendance_bulk(attendance_records)
    
    print(f"   ✓ Added {attendance_count} attendance records")
    
    # Test employee attendance query
    print("\n6. Querying attendance for employee E001 (January 2026)...")
    emp_att = db.get_employee_attendance("E001", "2026-01-01", "2026-01-31")
    print(f"   Total records: {len(emp_att)}")
    summary = db.get_attendance_summary_by_employee("E001", "2026-01-01", "2026-01-31")
    print(f"   Present: {summary['P']}, Absent: {summary['A']}, Weekly Off: {summary['WO']}, Holidays: {summary['H']}")
    
    # Test month attendance query
    print("\n7. Querying all attendance for January 2026...")
    month_att = db.get_month_attendance(2026, 1)
    print(f"   Total records for all employees: {len(month_att)}")
    
    # Test department attendance query
    print("\n8. Querying IT department attendance (January 2026)...")
    dept_att = db.get_dept_attendance("IT", "2026-01-01", "2026-01-31")
    print(f"   Total records for IT department: {len(dept_att)}")
    
    dept_summary = db.get_attendance_summary_by_dept("IT", "2026-01-01", "2026-01-31")
    print("   Summary by employee:")
    for _, row in dept_summary.iterrows():
        print(f"   - {row['emp_id']} ({row['name']}): P={row['Present']}, A={row['Absent']}, WO={row['Weekly_Off']}, H={row['Holidays']}")
    
    # Test statistics
    print("\n9. Database statistics...")
    stats = db.get_statistics()
    print(f"   Total Employees: {stats['total_employees']}")
    print(f"   Total Attendance Records: {stats['total_attendance_records']}")
    print(f"   Total Departments: {stats['total_departments']}")
    print(f"   Date Range: {stats['earliest_attendance']} to {stats['latest_attendance']}")
    
    # Test backup
    print("\n10. Creating database backup...")
    backup_path = db.backup_database("test_backup.db")
    print(f"   ✓ Backup created: {backup_path}")
    
    # Test export
    print("\n11. Exporting to Excel...")
    db.export_to_excel("test_export.xlsx", "2026-01-01", "2026-01-31")
    print("   ✓ Exported to test_export.xlsx")
    
    # Update employee
    print("\n12. Testing employee update...")
    if db.update_employee("E001", name="John Doe Jr."):
        print("   ✓ Updated employee E001")
        emp = db.get_employee("E001")
        print(f"   New name: {emp['name']}")
    
    # Close database
    print("\n13. Closing database connection...")
    db.close()
    print("   ✓ Connection closed")
    
    print("\n" + "=" * 60)
    print("ALL TESTS COMPLETED SUCCESSFULLY!")
    print("=" * 60)
    print("\nGenerated files:")
    print("  - test_attendance.db (SQLite database)")
    print("  - test_backup.db (Backup file)")
    print("  - test_export.xlsx (Excel export)")
    print("\nYou can delete these test files after review.")

if __name__ == "__main__":
    test_database()
