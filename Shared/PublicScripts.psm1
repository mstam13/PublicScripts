#Requires -Version 5.1
<#
.SYNOPSIS
    Shared utilities for PublicScripts repository scripts.
.DESCRIPTION
    Provides consistent timestamped logging and log rotation for all scripts
    in the PublicScripts repository.

    Exported functions
    ------------------
    Initialize-ScriptLog  Creates the log directory and initialises the module-internal
                          log file path.  Must be called before Write-ScriptLog.
    Write-ScriptLog       Writes a timestamped entry to the log file and the appropriate
                          PowerShell stream (INFO -> Write-Information, WARN -> Write-Warning,
                          ERROR -> Write-Error).  All file output uses UTF-8 without BOM,
                          compatible with PS 5.1 and PS 7+.
    Remove-OldLogs        Removes log files in a directory that are older than a configurable
                          retention window.  Supports -WhatIf / -Confirm.
#>

Set-StrictMode -Version Latest

# Module-scoped state — reset on every Import-Module -Force call.
$script:LogEncoding = [System.Text.UTF8Encoding]::new($false)   # UTF-8 without BOM
$script:LogFile     = $null

# ---------------------------------------------------------------------------

function Initialize-ScriptLog {
    <#
    .SYNOPSIS
        Creates the log directory and sets the module-internal log file path.
    .DESCRIPTION
        Must be called once at the start of a script, before any Write-ScriptLog calls.
        The log file name is: yyyyMMdd_HHmmss[_Tag]_ScriptName.log
    .PARAMETER LogDirectory
        Directory where the log file is written.  Created automatically when absent.
    .PARAMETER ScriptName
        Identifies the calling script in the log file name (e.g. 'Compare-GPOsByOU').
    .PARAMETER Tag
        Optional string inserted between the timestamp and the script name, e.g. a domain FQDN.
        Omit or pass an empty string to skip the tag segment.
    .OUTPUTS
        [string] — full path of the log file that will receive all Write-ScriptLog output.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $LogDirectory,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ScriptName,

        [Parameter()]
        [string] $Tag = ''
    )

    $null = New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction Stop

    $timestamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $fileName       = if ($Tag) {
        "${timestamp}_${Tag}_${ScriptName}.log"
    }
    else {
        "${timestamp}_${ScriptName}.log"
    }
    $script:LogFile = Join-Path $LogDirectory $fileName
    return $script:LogFile
}

# ---------------------------------------------------------------------------

function Write-ScriptLog {
    <#
    .SYNOPSIS
        Writes a timestamped log entry to the log file and the appropriate PS stream.
    .DESCRIPTION
        Requires Initialize-ScriptLog to have been called first.
        Level mapping:
          INFO  -> Write-Information (respects $InformationPreference; default Continue)
          WARN  -> Write-Warning
          ERROR -> Write-Error
        File output is UTF-8 without BOM (compatible with PS 5.1 and PS 7+).
    .PARAMETER Message
        The text to log.
    .PARAMETER Level
        Severity level: INFO (default), WARN, or ERROR.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string] $Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string] $Level = 'INFO'
    )

    if ($null -eq $script:LogFile) {
        throw 'Write-ScriptLog: log file not initialised — call Initialize-ScriptLog first.'
    }

    $entry = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [System.IO.File]::AppendAllLines($script:LogFile, [string[]] @($entry), $script:LogEncoding)

    if ($Level -eq 'ERROR')    { Write-Error   $Message }
    elseif ($Level -eq 'WARN') { Write-Warning $Message }
    else { Write-Information -MessageData $entry -InformationAction Continue }
}

# ---------------------------------------------------------------------------

function Remove-OldLog {
    <#
    .SYNOPSIS
        Removes log files older than the specified retention period.
    .DESCRIPTION
        Supports -WhatIf and -Confirm.
        Write-ScriptLog must be initialised before calling this function.
    .PARAMETER LogDirectory
        Directory containing log files to rotate.
    .PARAMETER Filter
        Wildcard filter passed to Get-ChildItem.  Defaults to '*.log'.
    .PARAMETER RetentionDays
        Files whose LastWriteTime is older than this many days are deleted.
        Valid range: 1–3650.  Defaults to 30.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $LogDirectory,

        [Parameter()]
        [string] $Filter = '*.log',

        [Parameter()]
        [ValidateRange(1, 3650)]
        [int] $RetentionDays = 30
    )

    if (-not (Test-Path -LiteralPath $LogDirectory)) { return }

    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $old    = Get-ChildItem -LiteralPath $LogDirectory -Filter $Filter -File -ErrorAction SilentlyContinue |
              Where-Object { $_.LastWriteTime -lt $cutoff }

    foreach ($file in $old) {
        if ($PSCmdlet.ShouldProcess($file.FullName, "Remove log file (older than $RetentionDays days)")) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-ScriptLog "Removed old log: '$($file.Name)'."
            }
            catch {
                Write-ScriptLog "Could not remove '$($file.Name)': $_" -Level WARN
            }
        }
    }

    if ($old.Count -gt 0) {
        Write-ScriptLog "Log rotation complete — $($old.Count) file(s) older than $RetentionDays days removed."
    }
}
