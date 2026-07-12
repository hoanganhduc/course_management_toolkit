API Reference
=============

.. automodule:: course_hoanganhduc.core
   :members:

.. automodule:: course_hoanganhduc.config
   :members:

.. automodule:: course_hoanganhduc.canvas
   :members:

.. automodule:: course_hoanganhduc.google_classroom
   :members:

Agent and Classroom50 modules
-----------------------------

These modules support the restricted agent surfaces and Classroom50 adapter.
Prefer the CLI entrypoints documented in :doc:`usage` over calling them as a
public Python API unless you are extending the toolkit.

- ``course_hoanganhduc.classroom50`` — agent-safe Classroom50 facade
- ``course_hoanganhduc.c50_agent`` / ``c50_flags`` / ``c50_ops`` / ``c50_cli``
- ``course_hoanganhduc.canvas_agent``
- ``course_hoanganhduc.gclass_agent``
- ``course_hoanganhduc.db_agent``
- ``course_hoanganhduc.course_agent_common`` — agent mode + allowlist helpers
