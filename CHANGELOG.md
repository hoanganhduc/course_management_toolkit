# Changelog

## [0.1.0] - 2025-12-27
- Packaged the original script as a Python CLI with install helpers and standardized flags.
- Added course-scoped config/credential storage with cached course codes.
- Rebuilt the no-args menu with sections, arrow-key navigation, and numeric quick-jump.
- Added `--clear-config` and `--clear-credentials` helpers plus Windows/Linux compatibility fixes.
- Canvas/Google Classroom sync now resolves duplicates; Canvas grade sync stores final scores only.
- Added override grades (flexible headers/aliases) with clearer reporting in MAT exports.
- Added AI model testing/listing, rate-limit fallback, and model details in reports.
- Expanded submission quality checks with configurable thresholds and richer diagnostics.
- Updated docs and samples; added GPL-3.0-only license.
