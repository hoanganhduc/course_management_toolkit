@echo off
setlocal
set VENV_PATH=%USERPROFILE%\.course_venv

REM make.bat install:
REM - Creates (or reuses) %USERPROFILE%\.course_venv
REM - Installs the package in editable mode
if /I "%1"=="clean" goto :clean
if /I "%1"=="install" goto :install

echo Usage: make.bat ^<install^|clean^>
exit /b 1

:install
if not exist "%VENV_PATH%" python -m venv "%VENV_PATH%"
"%VENV_PATH%\Scripts\python.exe" -m pip install --upgrade pip
"%VENV_PATH%\Scripts\pip.exe" install -r requirements.txt
"%VENV_PATH%\Scripts\pip.exe" install -e .
exit /b %ERRORLEVEL%

:clean
for /d /r %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d"
for /d /r %%d in (*.egg-info) do @if exist "%%d" rd /s /q "%%d"
if exist build rd /s /q build
if exist dist rd /s /q dist
if exist .pytest_cache rd /s /q .pytest_cache
if exist .mypy_cache rd /s /q .mypy_cache
if exist .ruff_cache rd /s /q .ruff_cache
if exist final_evaluations rd /s /q final_evaluations
if exist weekly_reports rd /s /q weekly_reports
for /d %%d in (_calendar_test*) do @if exist "%%d" rd /s /q "%%d"
del /q run_report.txt 2>nul
del /q final_grade_distribution.txt 2>nul
del /q emails_and_names.txt 2>nul
del /q classroom_roster.csv 2>nul
del /q course_calendar*.txt 2>nul
del /q course_calendar*.md 2>nul
del /q course_calendar*.ics 2>nul
del /q _calendar_test_with_makeup.txt 2>nul
del /q grade_diff_*.csv 2>nul
del /q *_updated.xlsx 2>nul
exit /b 0

