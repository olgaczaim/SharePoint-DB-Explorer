# Security policy

## Supported versions

Security fixes are made for the newest release in the [`releases`](releases) folder.

## Reporting a vulnerability

Use the repository's **Security > Report a vulnerability** option when private vulnerability reporting is enabled. If that option is unavailable, contact the maintainer privately through a contact method listed on their GitHub profile. Keep vulnerability details and source data out of public issues.

Include a description of the problem, the steps to reproduce it, and its possible impact. Do not include real content databases, exported documents or credentials.

## How the app handles data

- It uses read-only SQL queries. Connection values are handled by `SqlConnectionStringBuilder`, query values use parameters, and generated column projections use validated database schema metadata.
- Connection settings are entered for each session. SQL Server passwords are kept in memory only and are never written to disk, logs or reports.
- Encryption and certificate trust are configurable in the connection dialog.
- Exported files and reports are written only to the folder or archive you choose, without overwriting existing files.
- The log file (`%LOCALAPPDATA%\SharePointExplorer\Logs`) records storage warnings with document identifiers and paths, but never document contents or credentials.
