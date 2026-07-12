Course Management Toolkit
=========================

Utilities for managing course rosters, grading, OCR extraction, Canvas/Google Classroom workflows,
and Classroom50 (foundation50) roster sync, plus agent-safe entrypoints for AI coding agents.
Work in progress, mainly designed for personal use but open-sourced for others to adapt. Code with the help of GitHub Copilot and ChatGPT Codex.
Last updated: |today|

Package information
-------------------

- Version: 0.1.4 (+ unreleased Classroom50 / agent entrypoints)
- GitHub: https://github.com/hoanganhduc/course_management_toolkit
- Donations:
  - Buy Me a Coffee: https://www.buymeacoffee.com/hoanganhduc
  - Ko-fi: https://ko-fi.com/hoanganhduc

Recent updates
--------------

- Classroom50 (foundation50): list/sync/export roster via ``gh teacher`` wrapper; human-only download.
- Agent-safe modules: ``c50_agent``, ``canvas_agent``, ``gclass_agent``, ``db_agent`` with allowlists.
- Canvas/Google Classroom sync now normalizes scores to a 10-point scale when possible.
- MAT*.xlsx roster imports ignore score columns (CC, GK, CK, totals).
- Student detail exports support selectable sort orders.
- Duplicate-name reporting supports Name, Google Classroom Display Name, and Canvas Display Name with TXT/CSV/JSON exports.

Contents
--------

.. toctree::
   :maxdepth: 2

   usage
   cli_reference
   api
   changelog
