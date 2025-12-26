# Changelog

## [0.1.0] - 2025-12-26
- Converted the original script into a Python package with CLI entry point.
- Added packaging metadata, CLI entry points, and install helpers for Windows/Linux/macOS.
- Added course code prompting with `.course_code` caching and course-scoped config/credentials/token paths.
- Standardized CLI flags (`--version`, consistent long flags, memorable shortcuts).
- Canvas/Google Classroom sync now resolves duplicates by priority and prompts on conflicts.
- Canvas grading sync now processes all assignments/quizzes and stores final scores only.
- Fixed CK/GK column detection to avoid attendance/quiz/assignment columns.
- Default student database path now uses the current working directory.
- Added override_grades.xlsx support to override CC/CK/GK with reasons in final evaluation output.
- Moved sample files into `sample/` with documentation in `sample/index.md`.
- Added `sample/credentials.sample.json` as a Google service account template.
