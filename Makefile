VENV_PATH ?= $(HOME)/.course_venv
PYTHON ?= python3

# make install:
# - Creates (or reuses) the per-user venv at ~/.course_venv.
# - Installs the package in editable mode.
install:
	@echo "Creating virtualenv at $(VENV_PATH)"
	@if [ ! -d "$(VENV_PATH)" ]; then $(PYTHON) -m venv $(VENV_PATH); fi
	@$(VENV_PATH)/bin/python -m pip install --upgrade pip
	@$(VENV_PATH)/bin/pip install -r requirements.txt
	@$(VENV_PATH)/bin/pip install -e .

clean:
	@find . -type d -name "__pycache__" -prune -exec rm -rf {} +
	@find . -type d -name "*.egg-info" -prune -exec rm -rf {} +
	@rm -rf build dist .pytest_cache .mypy_cache .ruff_cache
	@rm -rf final_evaluations weekly_reports _calendar_test*
	@rm -f run_report.txt final_grade_distribution.txt emails_and_names.txt classroom_roster.csv
	@rm -f course_calendar*.txt course_calendar*.md course_calendar*.ics
	@rm -f _calendar_test_with_makeup.txt grade_diff_*.csv *_updated.xlsx
