import pandas as pd
import os

filename = "Danh sách công ty liên hệ (không xóa).xlsx"
file_path = os.path.join(os.getcwd(), filename)

try:
    df = pd.read_excel(file_path)
    print("Columns:")
    for col in df.columns:
        print(f"- {col}")
    print("\nFirst 3 rows:")
    print(df.head(3).to_string())
except Exception as e:
    print(f"Error reading excel: {e}")
