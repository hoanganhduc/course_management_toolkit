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
