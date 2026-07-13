# Shared/PublicScripts.psm1

> Version 1.0.0 — 2026-07-13

## Synopsis

Shared logging utilities used by all scripts in this repository.

## Description

`PublicScripts.psm1` is a PowerShell module that provides the logging infrastructure
shared by every script in this repository. Importing it once per script replaces the
per-script `Write-Log` / `Remove-OldLogs` boilerplate that had accumulated across
the collection.

The companion manifest `PublicScripts.psd1` declares compatible PowerShell versions
(5.1+) and the exported functions.

## Functions

### `Initialize-ScriptLog`

Creates the log directory if absent, builds the log-file name, and stores the path in
module scope so `Write-ScriptLog` can use it without requiring callers to pass it each
time.

| Parameter | Type | Mandatory | Default | Description |
| --- | --- | --- | --- | --- |
| `-LogDirectory` | String | Yes | — | Directory where log files are written. Created automatically if absent. |
| `-ScriptName` | String | Yes | — | Base name embedded in the log filename (no extension). |
| `-Tag` | String | No | — | Optional label inserted between the timestamp and the script name (e.g. a domain FQDN). |

**Returns** the full path to the log file (`[string]`).

**Log filename format:**

| Tag provided | Filename |
| --- | --- |
| No | `yyyyMMdd_HHmmss_<ScriptName>.log` |
| Yes | `yyyyMMdd_HHmmss_<Tag>_<ScriptName>.log` |

---

### `Write-ScriptLog`

Appends a timestamped entry to the log file opened by `Initialize-ScriptLog` and
echoes the message to the appropriate PowerShell output stream.

| Parameter | Type | Mandatory | Default | Description |
| --- | --- | --- | --- | --- |
| `-Message` | String | Yes | — | The text to record. Position 0 — positional calls supported. |
| `-Level` | String | No | `INFO` | Severity: `INFO`, `WARN`, or `ERROR`. |

**Stream routing:**

| Level | Stream |
| --- | --- |
| `INFO` | `Write-Information` (tag `LogEntry`) |
| `WARN` | `Write-Warning` |
| `ERROR` | `Write-Error` |

**Log entry format:** `[yyyy-MM-dd HH:mm:ss] [LEVEL] Message`

**Encoding:** UTF-8 without BOM via `[System.IO.File]::AppendAllLines`.

---

### `Remove-OldLog`

Deletes log files in a directory that are older than the specified retention period.
Supports `-WhatIf` and `-Confirm`.

| Parameter | Type | Mandatory | Default | Description |
| --- | --- | --- | --- | --- |
| `-LogDirectory` | String | Yes | — | Directory to scan for old log files. Returns silently if the directory does not exist. |
| `-Filter` | String | No | `*.log` | Wildcard filter passed to `Get-ChildItem`. Pass a script-specific pattern (e.g. `'*MyScript.log'`) to avoid touching other scripts' logs. |
| `-RetentionDays` | Int | No | `30` | Files whose `LastWriteTime` is older than this many days are deleted. Valid range: 1–3650. |

## Usage

```powershell
Import-Module (Join-Path $PSScriptRoot '..\Shared\PublicScripts.psm1') -Force

# Initialise — call once at script startup
$null = Initialize-ScriptLog -LogDirectory (Join-Path $PSScriptRoot 'Log') `
    -ScriptName 'MyScript' -Tag $Domain

Write-ScriptLog 'Script started.'
Write-ScriptLog -Level WARN 'Something looks suspicious.'
Write-ScriptLog 'Done.' -Level INFO

# Rotate old logs at script end
Remove-OldLog -LogDirectory (Join-Path $PSScriptRoot 'Log') `
    -Filter '*MyScript.log' -RetentionDays 30
```

## Notes

- `Write-ScriptLog` will throw if called before `Initialize-ScriptLog` (the internal
  `$script:LogFile` is `$null`).
- All scripts import with `-Force` so the module is reloaded cleanly on repeated dot-sourcing
  in interactive sessions.
- The module is not published to PSGallery; it is distributed alongside the repository
  scripts via the `Shared/` folder.

## Version History

| Version | Date | Author | Changes |
| --- | --- | --- | --- |
| 1.0.0 | 2026-07-13 | M. Stam | Initial release. Extracted `Initialize-ScriptLog`, `Write-ScriptLog`, and `Remove-OldLog` from per-script boilerplate across all five scripts in the repository. |
