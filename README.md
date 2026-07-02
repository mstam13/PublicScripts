# PublicScripts

A collection of reusable PowerShell and T-SQL scripts for enterprise infrastructure management.

## Contents

| Folder | Version | Type | Description |
| --- | --- | --- | --- |
| [Add-Named-Instances](Add-Named-Instances/) | 1.9.0 | PowerShell | Assigns SQL Server named instance service accounts to the required local security groups for CIS hardening. Creates missing groups automatically; verifies LSA rights via `secedit`. Supports `-WhatIf`. |
| [CleanupPolicies](CleanupPolicies/) | 1.2.0 | PowerShell | Inventories Group Policy Objects that are unlinked (not applied to any OU, site or domain container) or have no 'Apply Group Policy' Allow ACE, producing an Excel/CSV report for GPO cleanup reviews. |
| [Compare-GPOsByOU](Compare-GPOsByOU/) | 1.0.0 | PowerShell | Reports all GPOs linked to each OU including link order, enabled/enforced state, and full WMI filter details (name, description, WQL query) for side-by-side GPO comparison. |
| [Generate-RDCManConfigs](Generate-RDCManConfigs/) | 2.0.0 | PowerShell | Generates RDCMan (`.rdg`) configuration files for one or more Active Directory domains by querying enabled Windows Server objects and organising them into OU-mirrored groups. |
| [Get-ServiceAccounts](Get-ServiceAccounts/) | 1.7 | PowerShell | Scans all enabled Windows Server objects in Active Directory for non-standard accounts used by Windows services and scheduled tasks, exporting results to Excel and per-server CSV files. |
| [SQL-Create-ServiceNow-User](SQL-Create-ServiceNow-User/) | 1.0 | T-SQL | Provisions a Windows-authenticated SQL Server login and database user for the local `<MachineName>\servicenow` account, used by ServiceNow discovery and integration. |

## Usage

Each folder contains:

- A `.sql` or `.ps1` script — the main executable.
- A `.md` file — full documentation (synopsis, prerequisites, usage, notes).
- Additional supporting files where applicable (`.mmd` diagrams, sample config, etc.).

Refer to the `.md` file in each folder for detailed instructions before running any script.

## Requirements

- **SQL scripts** — SQL Server 2016 or later; SSMS or `sqlcmd` for execution.
- **PowerShell scripts** — PowerShell 5.1 or PowerShell 7+; modules listed per script.
  - `Add-Named-Instances` requires local administrator privileges (`#Requires -RunAsAdministrator`); no additional modules.
  - `CleanupPolicies` requires the `GroupPolicy` and `ActiveDirectory` modules (RSAT); optionally `ImportExcel` for `.xlsx` output.
  - `Compare-GPOsByOU` requires the `GroupPolicy` and `ActiveDirectory` modules (RSAT); optionally `ImportExcel` for `.xlsx` output.
  - `Generate-RDCManConfigs` requires the `ActiveDirectory` module (RSAT).
  - `Get-ServiceAccounts` requires the `ActiveDirectory` module (RSAT) and `ImportExcel`; WMI/CIM access to target servers (WinRM preferred, DCOM fallback).

## License

[MIT](LICENSE)
