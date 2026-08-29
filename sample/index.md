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

Gradebook and Company samples (`sample/mat/`, `sample/overrides/`, `sample/`):
- `MAT-examples.xlsx`: Sample course grade sheet used at VNU University of Science (Hanoi), with the original header/footer layout preserved, 10 placeholder students, and CC/GK/CK columns filled with sample values.
- `override_grades.xlsx`: Example override file. Columns required: Ma Sinh Viˆn or H? v… Tˆn, plus at least one of CC/GK/CK (order does not matter). `STT` and `Ly do` are optional. Common header aliases are accepted (for example `MSSV`, `H? tˆn`, `Midterm`, `Final`, `Reason`). Non-empty CC/GK/CK cells replace computed grades; Ly do explains why.
- `companies_sample.csv`: Sample companies data in English for testing imports.

Config samples (`sample/config/`):
- `config.sample.json`: Full configuration template for local setup.
- `credentials.sample.json`: Google service account credential template for legacy
  workflows; service accounts are not accepted by the human assignment creator.

Google Classroom assignment samples (`sample/google_classroom/`):
- `assignment-minimal.sample.json`: Draft assignment with no attachment or rubric.
- `assignment-test-draft.sample.json`: Safe no-attachment draft used for a live
  integration check.
- `assignment-full.sample.json`: Every stable supported option. Replace all IDs and
  the example timestamps before previewing or creating it.

All placeholder student names/IDs are consistent across the files above.
