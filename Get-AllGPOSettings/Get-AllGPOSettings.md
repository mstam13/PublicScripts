# Get-AllGPOSettings.ps1

Reports every configured setting in every Group Policy Object in a domain,
regardless of where — or whether — each GPO is linked. Produces a single
flat "Settings" report suitable for full-domain policy review, auditing, or
change tracking.

## Synopsis

```powershell
.\Get-AllGPOSettings.ps1 [-Domain <string>] [-DomainController <string>]
                          [-OutputPath <string>] [-PassThru]
```

## Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `Domain` | String | No | `$env:USERDNSDOMAIN` | FQDN of the AD domain to query |
| `DomainController` | String | No | — (Windows selects) | FQDN or hostname of a specific DC; pins all GPO queries to one source to avoid replication-lag inconsistencies |
| `OutputPath` | String | No | Script directory | Directory where the report is written; created automatically if it does not exist |
| `PassThru` | Switch | No | — | Return the collected settings as an array of objects to the pipeline |

## Output

| File | When | Description |
| --- | --- | --- |
| `YYYY-MM-dd_<Domain>_Get-AllGPOSettings.xlsx` | ImportExcel available | Single-worksheet workbook |
| `YYYY-MM-dd_<Domain>_Get-AllGPOSettings.csv` | Fallback | One row per configured setting |
| `Log\YYYYMMDD_HHmmss_<Domain>_Get-AllGPOSettings.log` | Always | Timestamped run log |

### Settings worksheet columns

These mirror the `Settings` worksheet produced by
[Compare-GPOsByOU.ps1](../Compare-GPOsByOU/Compare-GPOsByOU.ps1) with
`-CompareSettings`, minus the OU-link-specific columns (`ContainerName`,
`ContainerDN`, `LinkOrder`, `LinkEnabled`, `LinkEnforced`) — which have no
meaning here since a row is no longer tied to a single OU link. `GPOId` and
`GPOStatus` are added instead, so each row can be traced back to an
unambiguous GPO even when display names collide.

| Column | Description |
| --- | --- |
| `GPOName` | Display name of the GPO |
| `GPOId` | GUID in `{…}` notation |
| `GPOStatus` | `AllSettingsEnabled` / `UserSettingsDisabled` / `ComputerSettingsDisabled` / `AllSettingsDisabled` |
| `WMIFilterName` | Attached WMI filter (blank = none) |
| `ComputerSettingsEnabled` | Whether Computer Configuration is active on the GPO |
| `UserSettingsEnabled` | Whether User Configuration is active on the GPO |
| `Area` | `Computer` or `User` |
| `ExtensionType` | `Administrative Templates`, `Security Settings`, `Scripts`, or `Windows Firewall` |
| `Category` | Sub-category within the extension type |
| `SettingName` | Name of the configured setting |
| `SettingState` | State (e.g. `Enabled`, `Disabled`, `Configured`) |
| `SettingValue` | Configured value |

## Requirements

| Requirement | Details |
| --- | --- |
| PowerShell | 5.1 or 7+ |
| RSAT — Group Policy | `GroupPolicy` module (`RSAT-GPMC`) |
| ImportExcel | Optional — [dfinke/ImportExcel](https://github.com/dfinke/ImportExcel). Falls back to CSV when not installed. |
| Permissions | Group Policy Read on all GPOs |

Install RSAT feature (elevated, Windows 10/11 / Server 2016+):

```powershell
Add-WindowsCapability -Online -Name 'Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0'
```

Install ImportExcel:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser
```

## Examples

### Run against the current user's domain

```powershell
.\Get-AllGPOSettings.ps1
```

### Run against a specific domain and write output to a custom folder

```powershell
.\Get-AllGPOSettings.ps1 -Domain contoso.com -OutputPath C:\Reports
```

### Pin all queries to a specific domain controller

```powershell
.\Get-AllGPOSettings.ps1 -Domain contoso.com -DomainController dc01.contoso.com
```

### Capture the results for further filtering in the pipeline

```powershell
$settings = .\Get-AllGPOSettings.ps1 -PassThru
$settings | Where-Object { $_.ExtensionType -eq 'Administrative Templates' }
```

## Notes

- Each GPO report requires one network round-trip to the domain controller.
  On domains with many GPOs this may significantly increase runtime.
- On PowerShell 7+, reports are fetched in parallel (up to 8 concurrent
  requests) for faster execution.
- Orphaned (unlinked) GPOs are included, since this report is scoped to
  GPOs rather than to OU links. Use
  [Compare-GPOsByOU.ps1](../Compare-GPOsByOU/Compare-GPOsByOU.ps1) if you
  need settings reported per OU link instead, including conflict detection
  across GPOs linked to the same container.

### Version history

| Version | Date | Author | Changes |
| --- | --- | --- | --- |
| 1.0.1 | 2026-08-20 | M. Stam | Fixed a crash ("The property 'Name' cannot be found on this object") that occurred when a GPO referenced a WMI filter deleted from Active Directory; WMI filter name resolution now uses try/catch instead of a bare null-check. |
| 1.0.0 | 2026-08-20 | M. Stam | Initial release. |
