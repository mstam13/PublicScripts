# PublicScripts

A collection of reusable PowerShell and T-SQL scripts for enterprise infrastructure management.

## Contents

| Folder | Type | Description |
| --- | --- | --- |
| [Generate-RDCManConfigs](Generate-RDCManConfigs/) | PowerShell | Generates RDCMan (`.rdg`) configuration files for one or more Active Directory domains by querying enabled Windows Server objects and organising them into OU-mirrored groups. |
| [CleanupPolicies](CleanupPolicies/) | PowerShell | Inventories Group Policy Objects that are unlinked (not applied to any OU, site or domain container) or have no 'Apply Group Policy' Allow ACE, producing an Excel/CSV report for GPO cleanup reviews. |
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
  - `Generate-RDCManConfigs` requires the `ActiveDirectory` module (RSAT).
  - `CleanupPolicies` requires the `GroupPolicy` and `ActiveDirectory` modules (RSAT); optionally `ImportExcel` for `.xlsx` output.

## License

[MIT](LICENSE)
