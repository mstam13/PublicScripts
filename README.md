# PublicScripts

A collection of reusable PowerShell and T-SQL scripts for enterprise infrastructure management.

## Contents

| Folder | Type | Description |
| --- | --- | --- |
| [SQL-Create-ServiceNow-User](SQL-Create-ServiceNow-User/) | T-SQL | Provisions a Windows-authenticated SQL Server login and database user for the local `<MachineName>\servicenow` account, used by ServiceNow discovery and integration. |

## Usage

Each folder contains:

- A `.sql` or `.ps1` script — the main executable.
- A `.md` file — full documentation (synopsis, prerequisites, usage, notes).
- Additional supporting files where applicable (`.mmd` diagrams, sample config, etc.).

Refer to the `.md` file in each folder for detailed instructions before running any script.

## Requirements

- **SQL scripts** — SQL Server 2016 or later; SSMS or `sqlcmd` for execution.
- **PowerShell scripts** — PowerShell 5.1 or PowerShell 7+; modules listed per script.

## License

[MIT](LICENSE)
