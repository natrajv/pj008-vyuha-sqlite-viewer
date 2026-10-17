# Vyuha SQLite Viewer

## Features

* Open and remember a SQLite database file.
* Browse tables with wildcard table filtering and pagination.
* Search grid rows by wildcard or regular expression.
* Sort by clicking column headers.
* Run and save read-only SQL queries.
* Reorder or delete named queries.
* Export selected tables with data to a SQLite database or an Excel workbook, with one worksheet per table.
* Export selected tables without data to a blank SQLite database or Excel workbook, with column names on each worksheet.
* Search the export table list with * and ? wildcards; select all, deselect all, or invert the selection.
* Review selected tables and confirm the format before exporting; data exports require at least one table.
* Read-only database connections protect the source data.
 
## v1.4.2-20261018_0440
* Added new feature: Export selected tables with data to a SQLite database or an Excel workbook, with one worksheet per table.
* Added new feature: Export selected tables without data to a blank SQLite database or Excel workbook, with column names on each worksheet.

## v1.3.4 
* Saved Queries wildcard search (`*`, `?`) with Search and Clear.
* Removed Search Query Results from SQL tab.
* User-defined Saved Queries groups; select a group from the pulldown to restore its query list.
* Group names and queries persist in config.json.
* Startup database-not-found message clears when a database is successfully opened.
* Extract the ZIP and run `run.bat`.