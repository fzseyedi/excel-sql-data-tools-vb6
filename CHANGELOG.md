# Changelog

All notable changes to this project will be documented in this file.

## [1.2.0] - 2026-04-27

### Added

- Added a new main form (`frmMain`) to manage shared SQL Server connection, database selection, table selection, and tool launching.
- Added a new Export tool form (`frmExportSqlToExcel`).
- Added support for exporting one or more SQL Server tables.
- Added checkbox-based table selection for export.
- Added checkbox-based field selection per selected table.
- Added default table selection in Export form based on the table selected in the main form.
- Added table search support in the Export form.
- Added right-click context menu for selecting all or none in table and field lists.
- Added CSV export support.
- Added Excel export support:
  - `.xlsx` export using Excel Automation when Microsoft Excel is installed.
  - `.xls` export using ADO/OLEDB when Excel is not installed but the required provider is available.
- Added export output folder selection.
- Added overwrite confirmation for existing output files.
- Added export progress tracking.
- Added export cancel support.

### Changed

- Refactored the Import form to use the shared application context from `frmMain`.
- Moved SQL Server connection, database selection, and table selection responsibilities out of the Import form.
- Improved separation between shared application state and Import-specific workflow.
- Improved Import summary counts so staging and final target import counts are no longer mixed.
- Improved Export UI behavior so field selections are preserved when switching or searching tables.
- Improved Export UI behavior so the Cancel button remains clickable during export.

### Fixed

- Fixed incorrect Import success count where staging and target import counts were combined.
- Fixed field selection reset issues when searching and reselecting tables in the Export form.
- Fixed right-click “Select None” behavior for field selection.
- Fixed Export cancel button becoming unavailable during long-running export operations.
- Fixed UI container issue where disabling a parent frame prevented the Cancel button from being clicked.

### Notes

- CSV export is the most portable output option and does not require Microsoft Excel.
- Excel `.xlsx` export requires Microsoft Excel.
- Excel `.xls` export can work without Microsoft Excel when the required ADO/OLEDB provider is available.

## [1.1.0] - 2026-04-24

### Added
- Excel preview rows
- duplicate SQL-field mapping prevention
- stronger string length validation
- better numeric precision/scale validation
- improved NULL handling
- improved bit conversion handling
- cancel support during import
- optional save/load of last connection settings
- search support for databases and tables

### Improved
- clearer friendly error messages
- better preservation of original SQL/validation errors
- better duplicate error reporting
- better staging and target import error logging
- more responsive cancel handling using more frequent `DoEvents`
- better reset behavior when connection settings change
- better reset behavior when file/table selection changes
- improved progress/status behavior during import

### Fixed
- duplicate key errors being replaced by generic VB6 errors
- cancellation errors being overwritten during Excel Automation error handling
- stale database/table selections after failed reconnection or authentication change
- bit values being written incorrectly as `True/False` instead of `1/0`
- long Excel preview reads loading unnecessary full-row data

## [1.0.0] - 2026-04-24

### Added
- SQL Server connection with Windows Authentication
- SQL Server connection with SQL Authentication
- browsing non-system databases
- browsing base tables
- Excel reading using ADO/OLEDB
- Excel reading using Excel Automation
- Excel header loading
- manual mapping
- auto match
- save/load mapping
- required mapping validation
- staging table import
- final import into target table
- transaction handling
- rollback support
- `IDENTITY_INSERT` handling
- duplicate handling
- progress bar and execution status
- logging to file
- friendly error handling foundation
