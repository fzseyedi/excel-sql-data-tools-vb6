# VB6 Excel / SQL Import Export Tool

A practical **VB6 desktop utility** for importing Excel data into SQL Server and exporting SQL Server tables to Excel or CSV.

The application is designed for real-world database maintenance, data migration, and operational data exchange workflows, with support for shared SQL Server connection management, table/field selection, staging, validation, logging, progress tracking, and cancellation.

---

## Why this project?

Working with Excel and SQL Server is common in many business environments, but importing Excel data directly into production tables is risky.

This project was built to make import and export workflows safer, more controlled, and easier to use by providing:

- column mapping
- data validation
- staging tables before final import
- transaction-based final import
- rollback on failure
- progress tracking
- cancellation support
- detailed logging
- controlled SQL-to-Excel/CSV export

---

## Features

### Excel to SQL Server Import

- Import Excel data into SQL Server tables
- Preview Excel data before import
- Map Excel columns to SQL Server fields
- Save and load column mappings
- Validate required fields and data lengths
- Use staging tables before final import
- Support transaction-based final import workflow
- Roll back final import on failure
- Handle duplicate rows and validation errors
- Detailed import logging
- Progress tracking
- Cancel long-running import operations

### SQL Server to Excel / CSV Export

- Export one or more SQL Server tables
- Select fields separately for each selected table
- Search tables in the Export form
- Select all / none for tables and fields using right-click menu
- Export each selected table as a separate file
- CSV export support
- Excel `.xlsx` export using Excel Automation when Microsoft Excel is installed
- Excel `.xls` export using ADO/OLEDB when Excel is not installed but the required provider is available
- Output folder selection
- Overwrite confirmation for existing files
- Progress tracking during export
- Cancel long-running export operations

---

## Application Structure

The application now uses a simple main form as a shared shell for common SQL Server context.

### Main Form

The main form manages:

- SQL Server connection
- authentication mode
- database selection
- table selection
- launching Import and Export tools

### Import Form

The Import form focuses on:

- Excel file selection
- Excel preview
- column mapping
- mapping validation
- staging import
- final import
- progress and logging

### Export Form

The Export form focuses on:

- selecting one or more tables
- selecting fields per table
- choosing output type
- selecting output folder
- exporting to CSV, XLS, or XLSX
- progress and cancellation

---

## Import Flow

1. Connect to SQL Server from the main form
2. Select database and target table
3. Open the Import tool
4. Select Excel file
5. Load Excel columns
6. Preview Excel rows
7. Create or load column mappings
8. Validate mapping
9. Start import
10. Load rows into a staging table
11. Move rows from staging to target inside a transaction
12. Commit on success
13. Roll back on failure
14. Write detailed log output

---

## Export Flow

1. Connect to SQL Server from the main form
2. Select database and optionally a default table
3. Open the Export tool
4. Select one or more tables
5. Select fields for each table
6. Choose output type:
   - CSV
   - Excel
7. Select output folder
8. Start export
9. Confirm overwrite if output files already exist
10. Export each selected table as a separate file
11. Track progress and allow cancellation

---

## Safety Design

### Import Safety

This application does **not** insert Excel rows directly into the target table.

Instead, it uses:

- staging table first
- validation before final insert
- transaction for final import
- rollback on failure
- logging for troubleshooting

This reduces the risk of partial imports, invalid data, and inconsistent target tables.

### Export Safety

Export operations are read-only against SQL Server tables.

The Export tool:

- does not modify source tables
- asks before overwriting existing files
- keeps table and field selection explicit
- supports cancellation for long-running exports

---

## Current Status

This project is currently in a **working and testable state**.

Implemented and tested areas include:

- main form with shared SQL Server context
- connection and authentication
- database and table browsing
- Excel loading
- preview rows
- column mapping
- mapping save/load
- mapping validation
- staging import
- final transaction-based import
- rollback on failure
- duplicate handling
- import progress display
- import cancellation
- detailed import logging
- SQL Server table export
- multi-table export
- per-table field selection
- table search in Export form
- CSV export
- Excel `.xlsx` export with Excel Automation
- Excel `.xls` export with ADO/OLEDB fallback
- export progress display
- export cancellation
- improved error handling

---

## Tech Stack

- Visual Basic 6
- SQL Server
- ADO
- MSFlexGrid
- VB6 Common Controls
- Optional Excel Automation
- Optional ADO/OLEDB Excel provider support

---

## Requirements

Before running on another machine, verify that the following are available:

- Windows
- VB6 runtime
- ADO/MDAC
- required OCX controls used by the application
- SQL Server access
- Microsoft Excel for `.xlsx` export using Excel Automation
- Microsoft Jet/ACE OLEDB provider for `.xls` export using ADO/OLEDB

CSV export does **not** require Microsoft Excel.
For detailed installation and runtime dependency notes, see [INSTALL.md](INSTALL.md).

---

## Logs

The application writes log files into the `Logs` folder under the application path.

Import logs include information about:

- selected server and database
- target table
- Excel file
- staging table
- validation errors
- skipped rows
- duplicate errors
- final import status

---

## Roadmap

Planned or possible improvements include:

- richer final import/export summary
- separate row-level error log
- optional keep-staging-table mode
- better mapping visuals
- export filtering support
- saved export profiles
- deployment/setup checklist
- improved installer/package documentation

---

## Notes

This project is intended as a practical desktop data utility, not just a UI prototype.

It includes real import and export workflows with validation, staging, transaction control, logging, progress tracking, cancellation, and practical error handling.

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
