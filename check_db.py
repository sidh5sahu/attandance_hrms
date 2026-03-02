from database import AttendanceDatabase

db = AttendanceDatabase('attendance.db')
stats = db.get_statistics()

print("="*60)
print("DATABASE STATUS CHECK")
print("="*60)
print(f"\nTotal Employees: {stats['total_employees']}")
print(f"Total Attendance Records: {stats['total_attendance_records']}")
print(f"Departments: {stats['total_departments']}")
print(f"Date Range: {stats['earliest_attendance']} to {stats['latest_attendance']}")

if stats['total_employees'] == 0:
    print("\n⚠️ NO EMPLOYEES FOUND - You need to add employees first!")
    print("   Go to Employees tab → Add Employee or Import from Excel")

if stats['total_attendance_records'] == 0:
    print("\n⚠️ NO ATTENDANCE RECORDS - You need to upload attendance data!")
    print("   Go to Attendance tab → Upload Attendance File")
else:
    print(f"\n✅ Database has {stats['total_attendance_records']} attendance records")
    
    # Show sample records
    import pandas as pd
    query = "SELECT emp_id, date, punch_in, punch_out, status FROM attendance LIMIT 10"
    sample = pd.read_sql_query(query, db.conn)
    print("\nSample attendance records:")
    print(sample.to_string(index=False))

db.close()
print("\n" + "="*60)
