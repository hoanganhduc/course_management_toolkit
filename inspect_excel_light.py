import openpyxl
import os

filename = "Danh sách công ty liên hệ (không xóa).xlsx"
file_path = os.path.join(os.getcwd(), filename)

try:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    print("Columns (Row 1):")
    headers = []
    for cell in sheet[1]:
        headers.append(str(cell.value))
        print(f"- {cell.value}")
    print("\nRow 2:")
    for cell in sheet[2]:
        print(f"- {cell.value}")
except Exception as e:
    print(f"Error: {e}")
