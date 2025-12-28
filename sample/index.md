# Sample Files

This folder contains anonymized example inputs that mirror the expected formats.

- `MAT-examples.xlsx`: Sample course grade sheet used at VNU University of Science (Hanoi), with the original header/footer layout preserved, 10 placeholder students, and CC/GK/CK columns filled with sample values.
- `override_grades.xlsx`: Example override file. Columns required: Mã Sinh Viên or Họ và Tên, plus at least one of CC/GK/CK (order does not matter). `STT` and `Lý do` are optional. Common header aliases are accepted (for example `MSSV`, `Họ tên`, `Midterm`, `Final`, `Reason`). Non-empty CC/GK/CK cells replace computed grades; Lý do explains why.
- `config.sample.json`: Minimal configuration template for local setup.
- `credentials.sample.json`: Google service account credential template.

All placeholder student names/IDs are consistent across the files above.

