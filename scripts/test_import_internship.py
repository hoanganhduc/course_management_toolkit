import os
import sqlite3
import shutil
import tempfile
import csv
from course_hoanganhduc.data import import_internship_from_sheet

def test_local_import():
    # Create a dummy CSV file
    csv_content = [
        ['Dấu thời gian', 'Mã sinh viên', 'Họ và tên', 'Điện thoại', 'Email', 'Lớp (VD K66A4)', 'Môn học (Mã môn - Tên môn)', 'Công ty thực tập (Nếu công ty ngoài danh sách, các bạn gửi link trang web công ty bên cạnh tên công ty)', 'Bạn có là nhóm trưởng không? (do GV phân công)', 'Ngày bắt đầu thực tập', 'Hình thức thực tập', 'Link báo cáo thường xuyên'],
        ['21/01/2026 8:21:57', '12345678', 'Nguyen Van A', '0912345678', 'test@example.com', 'K67A2', 'MAT3371', 'Test Company', 'Khong', '01/02/2026', 'Fulltime', 'http://example.com/report']
    ]
    
    fd, csv_path = tempfile.mkstemp(suffix='.csv', text=True)
    os.close(fd)
    
    try:
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(csv_content)
            
        print(f"Created dummy CSV at: {csv_path}")
        
        # Create a temp DB
        fd_db, db_path = tempfile.mkstemp(suffix='.db')
        os.close(fd_db)
        
        print(f"Created temp DB at: {db_path}")
        
        # Run import
        import_internship_from_sheet(csv_path, db_path, verbose=True)
        
        # Verify DB content
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT student_id, full_name, internship_company FROM students WHERE student_id='12345678'")
        row = cursor.fetchone()
        conn.close()
        
        if row:
            print("Verification SUCCESS!")
            print(f"Found record: {row}")
            if row[0] == '12345678' and row[1] == 'Nguyen Van A' and row[2] == 'Test Company':
                print("Data matches expected values.")
            else:
                print("Data does NOT match expected values.")
        else:
            print("Verification FAILED! Record not found.")
            
    finally:
        if os.path.exists(csv_path):
            os.remove(csv_path)
        if os.path.exists(db_path):
            os.remove(db_path)

if __name__ == '__main__':
    test_local_import()
