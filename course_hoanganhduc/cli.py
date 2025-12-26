"""CLI entry point for the course management toolkit."""

from .core import main as core_main


def main():
    """Run the interactive course management CLI."""
    return core_main()


if __name__ == "__main__":
    raise SystemExit(main())
