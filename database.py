import sqlite3
import pandas as pd
from datetime import datetime
from typing import List, Dict, Optional, Tuple
import os

class AttendanceDatabase:
    """Database manager for employee attendance system"""
    
    def __init__(self, db_path: str = "attendance.db"):
        """Initialize database connection and create tables if they don't exist"""
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self.connect()
        self.create_tables()
    
    def connect(self):
        """Create database connection"""
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        # Enable foreign keys
        self.cursor.execute("PRAGMA foreign_keys = ON")
        self.conn.commit()
    
    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()
    
    def create_tables(self):
        """Create employees, attendance, and holidays tables"""
        # Create employees table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS employees (
                emp_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                dept TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Create attendance table with punch times and all leave types
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                emp_id TEXT NOT NULL,
                date DATE NOT NULL,
                punch_in TIME,
                punch_out TIME,
                status TEXT NOT NULL CHECK(status IN ('P', 'A', 'CL', 'EL', 'RH', 'HD', 'COFF', 'DL', 'WO', 'H')),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (emp_id) REFERENCES employees(emp_id) ON DELETE CASCADE,
                UNIQUE(emp_id, date)
            )
        ''')
        
        # Create holidays table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS holidays (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL UNIQUE,
                holiday_name TEXT NOT NULL,
                year INTEGER NOT NULL,
                type TEXT DEFAULT 'National',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Create indexes for faster queries
        self.cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_attendance_emp_id ON attendance(emp_id)
        ''')
        self.cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(date)
        ''')
        self.cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_employees_dept ON employees(dept)
        ''')
        self.cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_holidays_year ON holidays(year)
        ''')
        self.cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_holidays_date ON holidays(date)
        ''')
        
        self.conn.commit()
    
    # ==================== EMPLOYEE CRUD OPERATIONS ====================
    
    def add_employee(self, emp_id: str, name: str, dept: str) -> bool:
        """Add a new employee"""
        try:
            self.cursor.execute(
                "INSERT INTO employees (emp_id, name, dept) VALUES (?, ?, ?)",
                (emp_id, name, dept)
            )
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False  # Employee already exists
    
    def get_employee(self, emp_id: str) -> Optional[Dict]:
        """Get employee by ID"""
        self.cursor.execute(
            "SELECT emp_id, name, dept, created_at FROM employees WHERE emp_id = ?",
            (emp_id,)
        )
        row = self.cursor.fetchone()
        if row:
            return {
                'emp_id': row[0],
                'name': row[1],
                'dept': row[2],
                'created_at': row[3]
            }
        return None
    
    def get_all_employees(self) -> List[Dict]:
        """Get all employees"""
        self.cursor.execute("SELECT emp_id, name, dept, created_at FROM employees ORDER BY emp_id")
        rows = self.cursor.fetchall()
        return [
            {'emp_id': row[0], 'name': row[1], 'dept': row[2], 'created_at': row[3]}
            for row in rows
        ]
    
    def get_employees_by_dept(self, dept: str) -> List[Dict]:
        """Get all employees in a specific department"""
        self.cursor.execute(
            "SELECT emp_id, name, dept, created_at FROM employees WHERE dept = ? ORDER BY emp_id",
            (dept,)
        )
        rows = self.cursor.fetchall()
        return [
            {'emp_id': row[0], 'name': row[1], 'dept': row[2], 'created_at': row[3]}
            for row in rows
        ]
    
    def update_employee(self, emp_id: str, name: str = None, dept: str = None) -> bool:
        """Update employee information"""
        if name is None and dept is None:
            return False
        
        updates = []
        params = []
        if name is not None:
            updates.append("name = ?")
            params.append(name)
        if dept is not None:
            updates.append("dept = ?")
            params.append(dept)
        
        params.append(emp_id)
        query = f"UPDATE employees SET {', '.join(updates)} WHERE emp_id = ?"
        
        self.cursor.execute(query, params)
        self.conn.commit()
        return self.cursor.rowcount > 0
    
    def delete_employee(self, emp_id: str) -> bool:
        """Delete an employee (and all their attendance records due to CASCADE)"""
        self.cursor.execute("DELETE FROM employees WHERE emp_id = ?", (emp_id,))
        self.conn.commit()
        return self.cursor.rowcount > 0
    
    # ==================== ATTENDANCE OPERATIONS ====================
    
    def add_attendance(self, emp_id: str, date: str, status: str, punch_in: str = None, punch_out: str = None) -> bool:
        """Add or update attendance record for an employee on a specific date with punch times"""
        try:
            # Use INSERT OR REPLACE to handle duplicates
            self.cursor.execute(
                "INSERT OR REPLACE INTO attendance (emp_id, date, status, punch_in, punch_out) VALUES (?, ?, ?, ?, ?)",
                (emp_id, date, status, punch_in, punch_out)
            )
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def add_attendance_bulk(self, records: List[Tuple]) -> int:
        """Add multiple attendance records at once. 
        Records format: (emp_id, date, status) or (emp_id, date, status, punch_in, punch_out)
        Uses INSERT OR REPLACE - will replace existing records.
        Returns count of successful insertions."""
        count = 0
        for record in records:
            try:
                if len(record) == 3:
                    emp_id, date, status = record
                    punch_in, punch_out = None, None
                else:
                    emp_id, date, status, punch_in, punch_out = record
                
                self.cursor.execute(
                    "INSERT OR REPLACE INTO attendance (emp_id, date, status, punch_in, punch_out) VALUES (?, ?, ?, ?, ?)",
                    (emp_id, date, status, punch_in, punch_out)
                )
                count += 1
            except:
                continue
        self.conn.commit()
        return count
    
    def add_attendance_bulk_smart(self, records: List[Tuple]) -> Tuple[int, int]:
        """Smart bulk upload - only adds NEW records, skips existing ones.
        Records format: (emp_id, date, status) or (emp_id, date, status, punch_in, punch_out)
        Returns (new_records_added, duplicates_skipped)"""
        new_count = 0
        skip_count = 0
        
        for record in records:
            try:
                if len(record) == 3:
                    emp_id, date, status = record
                    punch_in, punch_out = None, None
                else:
                    emp_id, date, status, punch_in, punch_out = record
                
                # Check if record already exists
                self.cursor.execute(
                    "SELECT id FROM attendance WHERE emp_id = ? AND date = ?",
                    (emp_id, date)
                )
                
                if self.cursor.fetchone():
                    # Record exists, skip it
                    skip_count += 1
                else:
                    # New record, insert it
                    self.cursor.execute(
                        "INSERT INTO attendance (emp_id, date, status, punch_in, punch_out) VALUES (?, ?, ?, ?, ?)",
                        (emp_id, date, status, punch_in, punch_out)
                    )
                    new_count += 1
            except Exception as e:
                skip_count += 1
                continue
        
        self.conn.commit()
        return new_count, skip_count
    
    def get_attendance(self, emp_id: str, date: str) -> Optional[Dict]:
        """Get attendance details for a specific employee on a specific date"""
        self.cursor.execute(
            "SELECT status, punch_in, punch_out FROM attendance WHERE emp_id = ? AND date = ?",
            (emp_id, date)
        )
        row = self.cursor.fetchone()
        if row:
            return {
                'status': row[0],
                'punch_in': row[1],
                'punch_out': row[2]
            }
        return None
    
    def delete_attendance(self, emp_id: str, date: str) -> bool:
        """Delete attendance record"""
        self.cursor.execute(
            "DELETE FROM attendance WHERE emp_id = ? AND date = ?",
            (emp_id, date)
        )
        self.conn.commit()
        return self.cursor.rowcount > 0
    
    # ==================== HOLIDAY OPERATIONS ====================
    
    def add_holiday(self, date: str, holiday_name: str, year: int = None, hol_type: str = "National") -> bool:
        """Add a holiday to the database"""
        if year is None:
            year = int(date.split('-')[0])
        try:
            self.cursor.execute(
                "INSERT INTO holidays (date, holiday_name, year, type) VALUES (?, ?, ?, ?)",
                (date, holiday_name, year, hol_type)
            )
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def add_holidays_bulk(self, holidays: List[Tuple[str, str]]) -> int:
        """Add multiple holidays. Format: [(date, holiday_name), ...]
        Returns count of successful insertions."""
        count = 0
        for date, holiday_name in holidays:
            year = int(date.split('-')[0])
            if self.add_holiday(date, holiday_name, year):
                count += 1
        return count
    
    def get_holidays_by_year(self, year: int) -> List[Dict]:
        """Get all holidays for a specific year"""
        self.cursor.execute(
            "SELECT id, date, holiday_name, type FROM holidays WHERE year = ? ORDER BY date",
            (year,)
        )
        rows = self.cursor.fetchall()
        return [
            {'id': row[0], 'date': row[1], 'holiday_name': row[2], 'type': row[3]}
            for row in rows
        ]
    
    def get_all_holidays(self) -> List[Dict]:
        """Get all holidays"""
        self.cursor.execute("SELECT id, date, holiday_name, year, type FROM holidays ORDER BY date DESC")
        rows = self.cursor.fetchall()
        return [
            {'id': row[0], 'date': row[1], 'holiday_name': row[2], 'year': row[3], 'type': row[4]}
            for row in rows
        ]
    
    def delete_holiday(self, holiday_id: int) -> bool:
        """Delete a holiday"""
        self.cursor.execute("DELETE FROM holidays WHERE id = ?", (holiday_id,))
        self.conn.commit()
        return self.cursor.rowcount > 0
    
    def is_weekend(self, date_str: str) -> bool:
        """Check if a date is weekend (Saturday or Sunday)"""
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        return date_obj.weekday() in (5, 6)  # 5=Saturday, 6=Sunday
    
    def is_holiday(self, date_str: str) -> Optional[str]:
        """Check if a date is a holiday. Returns holiday name if true, None otherwise"""
        self.cursor.execute("SELECT holiday_name FROM holidays WHERE date = ?", (date_str,))
        row = self.cursor.fetchone()
        return row[0] if row else None
    
    def import_holidays_from_excel(self, file_path: str) -> Tuple[int, int]:
        """Import holidays from Excel file. 
        Expected columns: date, holiday_name, [type]
        Returns (success_count, error_count)"""
        try:
            df = pd.read_excel(file_path)
            if not all(col in df.columns for col in ['date', 'holiday_name']):
                raise ValueError("Excel file must have columns: date, holiday_name")
            
            # Convert date column
            df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')
            
            success = 0
            errors = 0
            for _, row in df.iterrows():
                date_str = row['date']
                holiday_name = row['holiday_name']
                year = int(date_str.split('-')[0])
                hol_type = row.get('type', 'National')
                
                if self.add_holiday(date_str, holiday_name, year, hol_type):
                    success += 1
                else:
                    errors += 1
            
            return success, errors
        except Exception as e:
            raise e
    
    # ==================== QUERY FUNCTIONS ====================
    
    def get_employee_attendance(self, emp_id: str, start_date: str = None, end_date: str = None) -> pd.DataFrame:
        """Get attendance records for a specific employee, optionally filtered by date range"""
        if start_date and end_date:
            query = """
                SELECT date, status FROM attendance 
                WHERE emp_id = ? AND date BETWEEN ? AND ?
                ORDER BY date
            """
            self.cursor.execute(query, (emp_id, start_date, end_date))
        else:
            query = "SELECT date, status FROM attendance WHERE emp_id = ? ORDER BY date"
            self.cursor.execute(query, (emp_id,))
        
        rows = self.cursor.fetchall()
        return pd.DataFrame(rows, columns=['date', 'status'])
    
    def get_month_attendance(self, year: int, month: int) -> pd.DataFrame:
        """Get attendance for all employees for a specific month"""
        start_date = f"{year}-{month:02d}-01"
        # Calculate last day of month
        if month == 12:
            end_date = f"{year}-12-31"
        else:
            end_date = f"{year}-{month:02d}-{pd.Period(f'{year}-{month:02d}').days_in_month}"
        
        query = """
            SELECT e.emp_id, e.name, e.dept, a.date, a.status
            FROM employees e
            LEFT JOIN attendance a ON e.emp_id = a.emp_id 
                AND a.date BETWEEN ? AND ?
            ORDER BY e.emp_id, a.date
        """
        self.cursor.execute(query, (start_date, end_date))
        rows = self.cursor.fetchall()
        return pd.DataFrame(rows, columns=['emp_id', 'name', 'dept', 'date', 'status'])
    
    def get_dept_attendance(self, dept: str, start_date: str = None, end_date: str = None) -> pd.DataFrame:
        """Get attendance for all employees in a department, optionally filtered by date range"""
        if start_date and end_date:
            query = """
                SELECT e.emp_id, e.name, e.dept, a.date, a.status
                FROM employees e
                LEFT JOIN attendance a ON e.emp_id = a.emp_id
                    AND a.date BETWEEN ? AND ?
                WHERE e.dept = ?
                ORDER BY e.emp_id, a.date
            """
            self.cursor.execute(query, (start_date, end_date, dept))
        else:
            query = """
                SELECT e.emp_id, e.name, e.dept, a.date, a.status
                FROM employees e
                LEFT JOIN attendance a ON e.emp_id = a.emp_id
                WHERE e.dept = ?
                ORDER BY e.emp_id, a.date
            """
            self.cursor.execute(query, (dept,))
        
        rows = self.cursor.fetchall()
        return pd.DataFrame(rows, columns=['emp_id', 'name', 'dept', 'date', 'status'])
    
    def get_attendance_summary_by_employee(self, emp_id: str, start_date: str = None, end_date: str = None) -> Dict:
        """Get attendance summary (counts of P, A, WO, H) for an employee"""
        if start_date and end_date:
            query = """
                SELECT status, COUNT(*) as count
                FROM attendance
                WHERE emp_id = ? AND date BETWEEN ? AND ?
                GROUP BY status
            """
            self.cursor.execute(query, (emp_id, start_date, end_date))
        else:
            query = """
                SELECT status, COUNT(*) as count
                FROM attendance
                WHERE emp_id = ?
                GROUP BY status
            """
            self.cursor.execute(query, (emp_id,))
        
        rows = self.cursor.fetchall()
        summary = {'P': 0, 'A': 0, 'WO': 0, 'H': 0}
        for status, count in rows:
            summary[status] = count
        return summary
    
    def get_attendance_summary_by_dept(self, dept: str, start_date: str = None, end_date: str = None) -> pd.DataFrame:
        """Get attendance summary for all employees in a department"""
        if start_date and end_date:
            query = """
                SELECT e.emp_id, e.name,
                    SUM(CASE WHEN a.status = 'P' THEN 1 ELSE 0 END) as Present,
                    SUM(CASE WHEN a.status = 'A' THEN 1 ELSE 0 END) as Absent,
                    SUM(CASE WHEN a.status = 'WO' THEN 1 ELSE 0 END) as Weekly_Off,
                    SUM(CASE WHEN a.status = 'H' THEN 1 ELSE 0 END) as Holidays
                FROM employees e
                LEFT JOIN attendance a ON e.emp_id = a.emp_id
                    AND a.date BETWEEN ? AND ?
                WHERE e.dept = ?
                GROUP BY e.emp_id, e.name
                ORDER BY e.emp_id
            """
            self.cursor.execute(query, (start_date, end_date, dept))
        else:
            query = """
                SELECT e.emp_id, e.name,
                    SUM(CASE WHEN a.status = 'P' THEN 1 ELSE 0 END) as Present,
                    SUM(CASE WHEN a.status = 'A' THEN 1 ELSE 0 END) as Absent,
                    SUM(CASE WHEN a.status = 'WO' THEN 1 ELSE 0 END) as Weekly_Off,
                    SUM(CASE WHEN a.status = 'H' THEN 1 ELSE 0 END) as Holidays
                FROM employees e
                LEFT JOIN attendance a ON e.emp_id = a.emp_id
                WHERE e.dept = ?
                GROUP BY e.emp_id, e.name
                ORDER BY e.emp_id
            """
            self.cursor.execute(query, (dept,))
        
        rows = self.cursor.fetchall()
        return pd.DataFrame(rows, columns=['emp_id', 'name', 'Present', 'Absent', 'Weekly_Off', 'Holidays'])
    
    # ==================== UTILITY FUNCTIONS ====================
    
    def get_all_departments(self) -> List[str]:
        """Get list of all unique departments"""
        self.cursor.execute("SELECT DISTINCT dept FROM employees ORDER BY dept")
        rows = self.cursor.fetchall()
        return [row[0] for row in rows]
    
    def get_statistics(self) -> Dict:
        """Get database statistics"""
        stats = {}
        
        # Total employees
        self.cursor.execute("SELECT COUNT(*) FROM employees")
        stats['total_employees'] = self.cursor.fetchone()[0]
        
        # Total attendance records
        self.cursor.execute("SELECT COUNT(*) FROM attendance")
        stats['total_attendance_records'] = self.cursor.fetchone()[0]
        
        # Total departments
        self.cursor.execute("SELECT COUNT(DISTINCT dept) FROM employees")
        stats['total_departments'] = self.cursor.fetchone()[0]
        
        # Date range of attendance records
        self.cursor.execute("SELECT MIN(date), MAX(date) FROM attendance")
        row = self.cursor.fetchone()
        stats['earliest_attendance'] = row[0]
        stats['latest_attendance'] = row[1]
        
        return stats
    
    def import_employees_from_excel(self, file_path: str) -> Tuple[int, int]:
        """Import employees from Excel file. Returns (success_count, error_count)"""
        try:
            df = pd.read_excel(file_path)
            if not all(col in df.columns for col in ['emp_id', 'name', 'Dept']):
                raise ValueError("Excel file must have columns: emp_id, name, Dept")
            
            success = 0
            errors = 0
            for _, row in df.iterrows():
                if self.add_employee(str(row['emp_id']), row['name'], row['Dept']):
                    success += 1
                else:
                    errors += 1
            
            return success, errors
        except Exception as e:
            raise e
    
    def import_attendance_from_excel(self, file_path: str) -> Tuple[int, int]:
        """Import attendance from Excel file. Returns (success_count, error_count)"""
        try:
            df = pd.read_excel(file_path)
            if not all(col in df.columns for col in ['emp_id', 'date']):
                raise ValueError("Excel file must have columns: emp_id, date")
            
            # Convert date column to standard format
            df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')
            
            records = [(str(row['emp_id']), row['date'], 'P') for _, row in df.iterrows()]
            success = self.add_attendance_bulk(records)
            errors = len(records) - success
            
            return success, errors
        except Exception as e:
            raise e
    
    def export_to_excel(self, output_path: str, start_date: str = None, end_date: str = None):
        """Export all data to Excel file with multiple sheets"""
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Export employees
            employees_df = pd.DataFrame(self.get_all_employees())
            employees_df.to_excel(writer, sheet_name='Employees', index=False)
            
            # Export attendance by department
            depts = self.get_all_departments()
            for dept in depts:
                dept_df = self.get_dept_attendance(dept, start_date, end_date)
                if not dept_df.empty:
                    sheet_name = str(dept)[:31]  # Excel sheet name limit
                    dept_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    def backup_database(self, backup_path: str = None):
        """Create a backup copy of the database"""
        if backup_path is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_path = f"attendance_backup_{timestamp}.db"
        
        import shutil
        shutil.copy(self.db_path, backup_path)
        return backup_path
