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
exit /b 0

