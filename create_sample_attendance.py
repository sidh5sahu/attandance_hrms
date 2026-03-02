"""
Create sample attendance data in EXACT 3-column format
Format: emp_id, date, time
"""
import pandas as pd
from datetime import datetime, timedelta

# Create sample attendance data
data = []
start_date = datetime(2026, 2, 1)

# Sample employees
employees = ["E001", "E002", "E003", "E004", "E005"]

# Generate attendance for February 2026 (first 10 days)
for day in range(10):
    current_date = start_date + timedelta(days=day)
    
    for emp_id in employees:
        # Multiple entries per day for punch time calculation
        
        # Entry 1: Morning punch in
        data.append({
            'emp_id': emp_id,
            'date': current_date.strftime('%Y-%m-%d'),
            'time': '09:00:00'
        })
        
        # Entry 2: Lunch out
        data.append({
            'emp_id': emp_id,
            'date': current_date.strftime('%Y-%m-%d'),
            'time': '13:00:00'
        })
        
        # Entry 3: Lunch in
        data.append({
            'emp_id': emp_id,
            'date': current_date.strftime('%Y-%m-%d'),
            'time': '14:00:00'
        })
        
        # Entry 4: Evening punch out
        data.append({
            'emp_id': emp_id,
            'date': current_date.strftime('%Y-%m-%d'),
            'time': '18:00:00'
        })

df = pd.DataFrame(data)

# Save with ONLY 3 columns
df = df[['emp_id', 'date', 'time']]
df.to_excel('sample_3col_attendance.xlsx', index=False)

print("="*60)
print("Created: sample_3col_attendance.xlsx")
print("="*60)
print(f"\nTotal entries: {len(df)}")
print(f"Unique records: {len(employees) * 10} (5 employees × 10 days)")
print(f"\nColumns (exactly 3):")
print("  1. emp_id")
print("  2. date")
print("  3. time (HH:MM:SS)")
print(f"\nFirst few rows:")
print(df.head(10))
print(f"\nSystem will calculate:")
print("  Punch In  = 09:00:00 (first entry per day)")
print("  Punch Out = 18:00:00 (last entry per day)")
print("\nWeekends automatically marked as WO")
print("="*60)
