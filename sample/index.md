# Sample Files

This folder contains anonymized example inputs that mirror the expected formats.

Calendar samples (`sample/calendar/`):
- `course_calendar_sample_input.txt`: Sample input that triggers a make-up week (holiday collision).
- `course_calendar_with_unofficial_holidays.txt`: Sample input that includes unofficial holidays.

Announcement samples (`sample/announcements/`):
- `announcement_input.txt`: Sample announcement input (Title/Message format).
- `announcement_refined_output.txt`: Sample announcement output after AI refinement.
- `announcement_input_vi.txt`: Sample announcement input in Vietnamese.
- `announcement_refined_output_vi.txt`: Sample announcement output in Vietnamese after AI refinement.

Gradebook samples (`sample/mat/`, `sample/overrides/`):
- `MAT-examples.xlsx`: Sample course grade sheet used at VNU University of Science (Hanoi), with the original header/footer layout preserved, 10 placeholder students, and CC/GK/CK columns filled with sample values.
- `override_grades.xlsx`: Example override file. Columns required: Ma Sinh Viˆn or H? v… Tˆn, plus at least one of CC/GK/CK (order does not matter). `STT` and `Ly do` are optional. Common header aliases are accepted (for example `MSSV`, `H? tˆn`, `Midterm`, `Final`, `Reason`). Non-empty CC/GK/CK cells replace computed grades; Ly do explains why.

Config samples (`sample/config/`):
- `config.sample.json`: Full configuration template for local setup.
- `credentials.sample.json`: Google service account credential template.

All placeholder student names/IDs are consistent across the files above.
