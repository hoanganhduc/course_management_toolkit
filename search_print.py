
import glob
import re

for filename in glob.glob('course_hoanganhduc/**/*.py', recursive=True):
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read()
            # Look for progress_report_data
            if "progress_report_data" in content:
                print(f"Match in {filename}")
                for i, line in enumerate(content.splitlines()):
                   if "progress_report_data" in line:
                       print(f"  Line {i+1}: {line.strip()}")
    except Exception as e:
        pass
