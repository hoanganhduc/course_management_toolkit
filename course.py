"""Backward-compatible entry point for running the CLI directly."""

from course_hoanganhduc.cli import main


if __name__ == "__main__":
    raise SystemExit(main())
