# Add-Named-Instances.ps1

> Version 1.9.0 — 2026-05-21

## Synopsis

Assigns SQL Server named instance service accounts to the required local security groups.

## Description

Discovers all SQL Server named instance services (SQL Server engine, SQL Agent, SSIS) and ensures their `NT Service\<ServiceName>` virtual accounts are members of the correct local security groups as required by CIS hardening guidelines.

The script is idempotent — groups that do not exist are created automatically, and accounts already present in a group are skipped. All operations support `-WhatIf` and `-Confirm`.

Every run writes a timestamped log file to the directory specified by `-LogDirectory` (default: `C:\Temp`). All actions — including successful additions, skipped accounts, and errors — are recorded there. When `-WhatIf` is specified the log is still written so the proposed changes are auditable. Log files older than 30 days are removed automatically at the end of each run.

The script exits with code **1** when any account/group operation fails, so GPO startup scripts and Task Scheduler can detect partial-hardening runs without inspecting the log.

After group memberships are updated, the effective local security policy is verified via `secedit /export`. For each target local group, the script resolves the group SID and checks whether it appears in the corresponding Windows User Rights Assignment privilege line in the exported policy. A `WARN` entry is written to the log for any gap — group membership alone has no security effect without the matching LSA right being explicitly assigned to that group.

## Parameters

| Parameter | Type | Default | Description |
| --- | --- | --- | --- |
| `-LogDirectory` | `string` | `C:\Temp` | Directory where timestamped log files are written. Created automatically if it does not exist. |
| `-WhatIf` | switch | — | Shows what would be changed without making any modifications. The log file is still written. |
| `-Confirm` | switch | — | Prompts for confirmation before each account/group addition. |
| `-Verbose` | switch | — | Writes per-account progress to the console in addition to the log file. |

## Group Mapping

| Service type | Local security group |
| --- | --- |
| SQL Server engine, SQL Agent | `AdjustMemoryQuotasForAProcess` |
| SQL Server engine, SQL Agent | `BypassTraverseChecking` |
| SQL Server engine, SQL Agent | `ReplaceProcessLevelToken` |
| SQL Server engine, SQL Agent | `LogonAsAService` |
| SSIS (`MsDtsServer*`) | `BypassTraverseChecking` |
| SSIS (`MsDtsServer*`) | `ImpersonateClientAfterAuthentication` |
| All NT Service virtual accounts (via Win32_Service) | `LogonAsAService` |

### Why these rights?

| Right | Reason |
| --- | --- |
| `AdjustMemoryQuotasForAProcess` | Allows SQL Server to manage memory allocations for its processes. |
| `BypassTraverseChecking` | Allows the service account to traverse directory trees without requiring explicit permissions on each folder. |
| `ReplaceProcessLevelToken` | Required for SQL Server and SQL Agent to start child processes under the service account identity. |
| `ImpersonateClientAfterAuthentication` | Required by SSIS to execute packages under the caller's identity after authentication. |
| `LogonAsAService` | Required for any virtual NT Service account to be registered as a service logon; without this right the service cannot start. |

## Requirements

- Must be run with **local administrator** privileges (`#Requires -RunAsAdministrator`). When run as a GPO startup script the process runs as SYSTEM, which satisfies this requirement.
- PowerShell 5.1 or later (uses `Get-LocalGroup`, `Get-LocalGroupMember`, `Add-LocalGroupMember`, `New-LocalGroup`).
- The machine execution policy must allow the script to run. Configure via GPO (`Computer Configuration > Windows Settings > Security Settings > Software Restriction Policies` or via `Turn on Script Execution` in Administrative Templates), or sign the script with a trusted code-signing certificate.

### GPO Startup Script configuration

Add a **PowerShell** startup script entry (not a legacy script entry) under
`Computer Configuration > Windows Settings > Scripts > Startup`:

| Field | Value |
| --- | --- |
| Script Name | `powershell.exe` |
| Script Parameters | `-NonInteractive -ExecutionPolicy RemoteSigned -File "\\<server>\<share>\Add-Named-Instances.ps1"` |

Or store the script locally on each machine (e.g. via a GPO file preference) and reference it as:

```text
-NonInteractive -ExecutionPolicy RemoteSigned -File "C:\Scripts\Add-Named-Instances.ps1"
```

> **Note:** `-NonInteractive` is required because startup scripts run without a user session. `-ExecutionPolicy RemoteSigned` overrides the machine policy for this invocation if needed; alternatively sign the script and use `AllSigned`.

## Usage

```powershell
# Dry run — shows what would be added without making changes (log is still written)
.\Add-Named-Instances.ps1 -WhatIf

# Run normally
.\Add-Named-Instances.ps1

# Write log files to a custom directory
.\Add-Named-Instances.ps1 -LogDirectory 'D:\Logs'

# Run with verbose per-account progress on the console
.\Add-Named-Instances.ps1 -Verbose

# Run with confirmation prompt for every action
.\Add-Named-Instances.ps1 -Confirm

# Capture results for reporting
$results = .\Add-Named-Instances.ps1
$results | Where-Object Action -eq 'Added'
$results | Export-Csv -Path ".\Add-Named-Instances-$(Get-Date -Format 'yyyy-MM-dd').csv" -NoTypeInformation
```

## Output

Each evaluated account/group combination emits a `PSCustomObject`:

| Property | Values | Description |
| --- | --- | --- |
| `Action` | `Added`, `WouldAdd`, `AlreadyMember`, `Failed` | Result of the operation. `WouldAdd` is emitted during a `-WhatIf` run. |
| `Account` | `NT Service\<ServiceName>` | The virtual service account. |
| `Group` | Group name string | The target local security group. |

## Logging

| Aspect | Detail |
| --- | --- |
| Location | `<LogDirectory>\Add-Named-Instances_yyyyMMdd_HHmmss.log` (default: `C:\Temp`) |
| Format | `[yyyy-MM-dd HH:mm:ss] [LEVEL] Message` |
| Levels | `INFO`, `WARN`, `ERROR` |
| Startup entry | Includes `$env:COMPUTERNAME` for machine-identifiable central log collection. |
| Retention | Log files older than **30 days** are deleted automatically at the end of each run, including runs where no instances are found. |

## Script Flow

```text
1. #region Functions                      — Write-Log, Remove-OldLogs, Add-AccountToGroup helpers
2. #region Configuration                  — Log file initialisation; group name constants
3. #region Discover                       — Enumerate SQL Server, SQL Agent, SSIS services
4. #region Ensure groups exist            — Create missing local security groups
5. #region Add members to groups          — Assign accounts to groups per CIS mapping
6.   #region Log cleanup                  — Remove log files older than 30 days
7. #region Verify LSA user-right assignments — secedit check: group SID vs privilege line
```

## LSA Verification

After memberships are set, the script calls `secedit /export` and parses the `[Privilege Rights]` section to verify each local group is wired to its expected Windows privilege:

| Local group | LSA privilege constant |
| --- | --- |
| `AdjustMemoryQuotasForAProcess` | `SeIncreaseQuotaPrivilege` |
| `BypassTraverseChecking` | `SeChangeNotifyPrivilege` |
| `ImpersonateClientAfterAuthentication` | `SeImpersonatePrivilege` |
| `ReplaceProcessLevelToken` | `SeAssignPrimaryTokenPrivilege` |
| `LogonAsAService` | `SeServiceLogonRight` |

For each group the script resolves its local SID and checks for `*<SID>` in the privilege line. Missing entries are logged as `[WARN] LSA GAP`. A `[INFO] LSA verification complete` summary line reports OK vs. gap counts.

> **Important:** this script only manages local group *membership*. The User Rights Assignment itself must be configured separately, either via a GPO (`Computer Configuration > Windows Settings > Security Settings > User Rights Assignment`) or via Local Security Policy (`secpol.msc`). Without that configuration, populating the groups has no security effect.

## Notes

- **Author:** Marcel Stam
- **Date:** 2026-05-21
- **Version:** 1.9.0
- Intended for use as part of CIS SQL Server hardening on Windows servers running named SQL Server instances.
- Service discovery uses `Get-Service -DisplayName '*SQL*'` and filters on display name patterns `SQL Server (*)` / `SQL Server Agent (*)` and service name pattern `MsDtsServer*`, so only **named instances** (not the default instance `MSSQLSERVER`) are targeted for the SQL engine/agent groups. The `LogonAsAService` assignment additionally covers all Windows services running as NT Service virtual accounts (via `Win32_Service`), with named-instance accounts deduplicated to prevent double-processing.
- Built-in system accounts (`LocalSystem`, `NT AUTHORITY\LocalService`, `NT AUTHORITY\NetworkService`) are excluded automatically from named-instance account lists.
- The script exits with code **1** when any `Failed` result is present; exit code **0** indicates full success.
- `Get-Service` and `Get-CimInstance` are wrapped in `try/catch` so discovery failures produce a friendly log entry before the script terminates.
- The LSA verification step is read-only and runs even under `-WhatIf`.
- All function parameters are decorated with `[Parameter(Mandatory)]` and `[ValidateNotNullOrEmpty()]` per coding guidelines.
- `$ErrorActionPreference` is set to `'Stop'` so unexpected errors outside `try/catch` are terminating rather than silently continuing in a non-interactive SYSTEM session.

## Version History

| Version | Date | Description |
| --- | --- | --- |
| 1.9.0 | 2026-05-21 | Removed unused `State` property from `$ServiceAccounts` query; switched to `Select-Object -ExpandProperty StartName` for a clean string list (improvement 6). Wrapped `Get-Service` and `Get-CimInstance` in `try/catch` for friendly error logging (improvement 7). Added `#region Verify LSA user-right assignments`: exports effective security policy via `secedit /export`, resolves each local group SID, and logs `WARN` for any group not referenced in its corresponding User Rights Assignment privilege line (improvement 8). |
| 1.8.0 | 2026-05-21 | Fixed stale `MemberCache` bug: `Add-AccountToGroup` now updates the in-memory cache after each successful addition, preventing spurious `Failed` results. Deduplicated `$ServiceAccounts` to exclude named-instance accounts already handled in earlier loops, eliminating the `LogonAsAService` double-processing (issues 1 & 2). Added `$env:COMPUTERNAME` to startup log entry (improvement 3). Made log directory configurable via `-LogDirectory` parameter with default `C:\Temp` (improvement 4). Added `exit 1` when `Failed` results are present (improvement 5). |
| 1.7.0 | 2026-05-20 | Compliance review: updated `.DESCRIPTION` to document `LogonAsAService` as a fifth target group for SQL engine/agent accounts and NT Service virtual accounts; corrected "four groups" → "five groups"; fixed misleading inline comment on `$LogonAsServiceGroup`; added missing `#endregion` closing `#region Add members to groups`. Updated markdown group-mapping and why-these-rights tables accordingly. |
| 1.6.0 | 2026-05-19 | GPO startup script compatibility: replaced em dash characters in `Write-Log` string literals with ` - ` to prevent a PS 5.1 cp1252 parse error (UTF-8 em dash byte `94` decodes to a curly-quote in cp1252, terminating string literals unexpectedly). Added `$ErrorActionPreference = 'Stop'` so unexpected errors outside `try/catch` are caught in a non-interactive SYSTEM session. |
| 1.5.0 | 2026-05-19 | Added `[Parameter()]` and `[ValidateNotNullOrEmpty()]` to all function parameters. Removed duplicate `Write-Error` from `Write-Log` ERROR level (log file entry is sufficient alongside `throw`). Fixed early-return path to run log cleanup before returning when no instances are found. |
| 1.4.0 | 2026-05-19 | Added `Remove-OldLogs` helper; log files older than 30 days are deleted from `C:\Temp` at the end of each run. |
| 1.3.0 | 2026-05-19 | Fixed `-WhatIf` bug: `Add-AccountToGroup` now emits a `WouldAdd` result instead of `$null`. Added `Write-Log` calls for `Added` and `AlreadyMember` actions. Summary now reports `WouldAdd` count. Expanded `.DESCRIPTION` and `.OUTPUTS`. |
| 1.2.0 | 2026-05-19 | Added `Write-Log` helper; timestamped log file written to `C:\Temp`; all verbose/warning output also persisted to the log file. |
| 1.1.0 | 2026-05-13 | Added Configuration region; early exit when no instances found; `MemberCache` parameter; `throw` instead of `exit 1`; result collection and verbose summary; explicit `param()` block. |
| 1.0.0 | 2026-05-13 | Initial release. |
