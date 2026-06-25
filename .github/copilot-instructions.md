# Workspace Guidelines

## Domain Context

This workspace contains PowerShell automation scripts for enterprise identity and infrastructure management at **TBI / SSC** (Service Delivery). Key domains:

- **Azure AD / Entra ID** — Microsoft Graph API, app registrations, certificate-based auth
- **On-premises Active Directory** — RSAT, nested group membership (`APL-*`, `RES-*`, `DEL-*` prefixes, domain: `contoso.com`, `global.contoso.com`)
- **Infrastructure** — SQL Server, Configuration Manager, DFS, Windows network administration
- **Language mix** — English code, Dutch variable names and comments are both acceptable

## Script Structure

Every script must include:

```powershell
<#
.SYNOPSIS   One-line summary
.DESCRIPTION  Full description
.PARAMETER  <Name>  Description per parameter
.OUTPUTS    Describe files/objects produced (paths, formats)
.EXAMPLE    Minimal working invocation
.NOTES      Author, date, CHG reference (if applicable)
#>
```

- Use `#region` / `#endregion` to divide logical sections
- Parameter blocks: `[Parameter(Mandatory=$true)]` + `[ValidateScript()]` or `[ValidateNotNullOrEmpty()]`
- Scheduled-task safe: use `$PSScriptRoot` / `$PSCommandPath` for relative paths, never `$PWD`

## Error Handling & Logging

- Wrap all external calls (AD, Graph, SQL) in `try/catch`; surface meaningful messages
- Log to `.\Log\` or `.\Logs\` with timestamp: `"$($script:LogFile) = "$PSScriptRoot\Log\$(Get-Date -Format 'yyyyMMdd_HHmmss').log"`
- Export reports to Excel via **ImportExcel** module (preferred); fall back to CSV when the module is unavailable

## Authentication

- **Certificate-based** for unattended / scheduled execution (Microsoft Graph); store `AppId`, `TenantId`, and `CertThumbprint` in a sidecar `.json` config, never hardcoded
- Support both interactive (`Connect-MgGraph -Scopes …`) and non-interactive (`Connect-MgGraph -ClientId … -Certificate …`) modes
- Document required Graph API permissions in the `.NOTES` or `.DESCRIPTION` block

## Common Modules

| Module | Purpose |
|--------|---------|
| `Microsoft.Graph.*` | Entra ID / Microsoft 365 operations |
| `ActiveDirectory` | On-prem AD (RSAT) |
| `ImportExcel` | Excel export without Office installed |
| `PSReadLine` | Interactive shell enhancements (profiles only) |

List `#Requires -Modules` at the top of each script where applicable.

## Naming & Output Conventions

- Verb-Noun cmdlet style: `Get-AplADGroupMembersNested`, `Set-BitLockerToAAD`
- Output files: `YYYY-MM-dd_ScriptName.xlsx` / `.csv`
- Config files: `ScriptName-config.json` alongside the script

## Architecture Notes

- `PublicScripts/` — one subfolder per script
