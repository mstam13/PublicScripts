<#
.SYNOPSIS
    Scans all AD servers for non-standard service and scheduled task accounts.

.DESCRIPTION
    Retrieves all Windows servers from Active Directory, checks whether each server
    is reachable, then collects accounts that are not standard built-ins from two sources:
      1. Windows services  (Win32_Service via CIM; DCOM fallback when WinRM is unavailable)
      2. Scheduled tasks   (Get-ScheduledTask via PSRemoting; skipped when WinRM is unavailable)
    Standard accounts excluded:
        LocalSystem, NT AUTHORITY\LocalService, NT AUTHORITY\NetworkService,
        NT AUTHORITY\*, NT SERVICE\*, well-known SIDs S-1-5-18/19/20,
        and empty / null identities.
    Results are exported to one CSV file per server in .\Output\ and combined into a
    single Excel workbook at the end of the run.
    All activity is written to a timestamped log file in .\Log\.

.PARAMETER SearchBase
    Optional. The distinguished name of the OU to search for server objects.
    Defaults to the entire domain.

.PARAMETER OutputFolder
    Optional. Path to the folder where per-server CSV files are written.
    Defaults to "$PSScriptRoot\Output".

.PARAMETER LogFolder
    Optional. Path to the folder where the log file is written.
    Defaults to "$PSScriptRoot\Log".

.PARAMETER ComputerList
    Optional. Path to a plain-text file (one server name per line) to scan instead
    of querying Active Directory. Lines starting with '#' and blank lines are ignored.
    Use the failed-servers file produced by a previous run to retry only those hosts.

.PARAMETER PingCount
    Optional. Number of ICMP echo requests sent to test reachability. Default: 1.

.OUTPUTS
    .\Output\<ServerName>_Services.csv                   — per-server services with non-standard accounts.
    .\Output\<ServerName>_ScheduledTasks.csv             — per-server scheduled tasks with non-standard accounts.
    .\Output\<timestamp>_ServiceAccounts.xlsx             — combined Excel workbook: 'Services' and 'ScheduledTasks' sheets.
    .\Output\<timestamp>_FailedServers.txt                — offline + error servers for retry.
    .\Log\<timestamp>_Get-ServiceAccounts.log             — full activity log.

.EXAMPLE
    .\Get-ServiceAccounts.ps1

.EXAMPLE
    .\Get-ServiceAccounts.ps1 -SearchBase "OU=Servers,DC=contoso,DC=com"

.EXAMPLE
    .\Get-ServiceAccounts.ps1 -ComputerList ".\Output\20260608_120000_FailedServers.txt"

.NOTES
    Author  : M. Stam
    Date    : 2026-06-10
    Requires: ActiveDirectory RSAT module, ImportExcel module, WMI/CIM access to target servers.
              PSRemoting (WinRM) is required for scheduled task enumeration; it is skipped
              gracefully when only DCOM is available.
              Run as an account with read access to AD and remote WMI/WinRM on all servers.
#>
#Requires -Modules ActiveDirectory, ImportExcel

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = "$PSScriptRoot\Output",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogFolder = "$PSScriptRoot\Log",

    [Parameter(Mandatory = $false)]
    [ValidateScript({ if ($_) { Test-Path $_ -PathType Leaf } else { $true } })]
    [string]$ComputerList,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 5)]
    [int]$PingCount = 1
)

#region --- Initialisation ---

# Failed servers (offline or CIM error) collected during this run
$script:FailedServers    = [System.Collections.Generic.List[string]]::new()
# CSV paths written during this run (prevents picking up files from previous runs)
$script:ServiceCsvPaths = [System.Collections.Generic.List[string]]::new()
$script:TaskCsvPaths    = [System.Collections.Generic.List[string]]::new()

# Standard / built-in service accounts to ignore
# Note: 'NT AUTHORITY\*' and 'NT SERVICE\*' wildcards already handle the prefixed forms;
# only bare 'localsystem' (no domain prefix) needs an explicit entry.
$StandardAccounts = @(
    'localsystem'
)

# Ensure output and log directories exist
foreach ($dir in $OutputFolder, $LogFolder) {
    if (-not (Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$timestamp             = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:LogFile        = Join-Path $LogFolder    "${timestamp}_Get-ServiceAccounts.log"
$script:FailedFile     = Join-Path $OutputFolder "${timestamp}_FailedServers.txt"
$script:ExcelFile      = Join-Path $OutputFolder "${timestamp}_ServiceAccounts.xlsx"

#endregion

#region --- Logging helper ---

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    $entry | Out-File -FilePath $script:LogFile -Append -Encoding utf8
    $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Gray' } }
    Write-Host $entry -ForegroundColor $color
}

#endregion

#region --- Retrieve servers from AD or file ---

Write-Log "Script started. Output: $OutputFolder | Log: $script:LogFile"

if ($PSBoundParameters.ContainsKey('ComputerList')) {
    Write-Log "Reading server list from file: $ComputerList"
    $serverNames = Get-Content -Path $ComputerList |
        Where-Object { $_ -notmatch '^\s*#' -and $_ -match '\S' } |
        ForEach-Object { $_.Trim() }
    # Build minimal objects so the rest of the loop works unchanged
    $servers = $serverNames | ForEach-Object {
        [PSCustomObject]@{ Name = $_; DNSHostName = $_ }
    }
    Write-Log "Loaded $($servers.Count) server(s) from file."
}
else {
    $adParams = @{
        Filter     = "OperatingSystem -like '*Windows Server*' -and Enabled -eq 'True'"
        Properties = 'Name', 'OperatingSystem', 'DNSHostName'
    }
    if ($PSBoundParameters.ContainsKey('SearchBase')) {
        $adParams['SearchBase'] = $SearchBase
    }

    Write-Log "Querying Active Directory for Windows Server objects..."
    try {
        $servers = Get-ADComputer @adParams
        Write-Log "Found $($servers.Count) enabled server(s) in AD."
    }
    catch {
        Write-Log "Failed to query Active Directory: $_" -Level ERROR
        exit 1
    }
}

#endregion

#region --- Scan each server ---

$totalOnline  = 0
$totalOffline = 0
$totalErrors  = 0
$serverIndex  = 0
$serverTotal  = @($servers).Count

foreach ($server in $servers) {
    $fqdn = if ($server.DNSHostName) { $server.DNSHostName } else { $server.Name }
    $serverIndex++
    Write-Progress -Activity 'Scanning servers' -Status "$fqdn ($serverIndex of $serverTotal)" -PercentComplete (($serverIndex / $serverTotal) * 100)

    #--- Reachability check ---
    Write-Log "Pinging $fqdn..."
    $online = Test-Connection -ComputerName $fqdn -Count $PingCount -Quiet -TimeoutSeconds 2 -ErrorAction SilentlyContinue

    if (-not $online) {
        Write-Log "$fqdn is OFFLINE or unreachable — skipping." -Level WARN
        $script:FailedServers.Add($server.Name)
        $totalOffline++
        continue
    }

    $totalOnline++
    Write-Log "$fqdn is online. Querying services and scheduled tasks..."

    #--- Query Win32_Service via CIM (falls back to DCOM if WSMan unavailable) ---
    try {
        $cimSession = $null
        $winRmWorks = $false

        # Try WSMan (WinRM) first; fall back to DCOM
        try {
            $cimSession = New-CimSession -ComputerName $fqdn -OperationTimeoutSec 30 -ErrorAction Stop
            $services   = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service -ErrorAction Stop
            $winRmWorks = $true
        }
        catch {
            Write-Log "$fqdn — WinRM unavailable, falling back to DCOM." -Level WARN
            if ($cimSession) { Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue }
            $dcomOption = New-CimSessionOption -Protocol Dcom
            $cimSession = New-CimSession -ComputerName $fqdn -SessionOption $dcomOption -OperationTimeoutSec 30 -ErrorAction Stop
            $services   = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service -ErrorAction Stop
        }

        #--- Filter non-standard service accounts ---
        $serviceRows = $services | Where-Object {
            $startName = ($_.StartName -replace '\\\\', '\').Trim()
            if ([string]::IsNullOrWhiteSpace($startName)) { return $false }
            if ($startName -like 'NT AUTHORITY\*')  { return $false }
            if ($startName -like 'NT SERVICE\*')    { return $false }
            if ($startName -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) { return $false }
            if ($StandardAccounts -contains $startName.ToLower()) { return $false }
            return $true
        } | Select-Object `
            @{N='Server';      E={ $server.Name }},
            @{N='FQDN';        E={ $fqdn }},
            @{N='Type';        E={ 'Service' }},
            Name,
            DisplayName,
            StartName,
            State,
            StartMode,
            PathName

        Write-Log "$fqdn — Services: $($services.Count) total | Non-standard accounts: $($serviceRows.Count)"

        #--- Query scheduled tasks via PSRemoting (requires WinRM) ---
        $taskRows = @()
        if ($winRmWorks) {
            try {
                $rawTasks = Invoke-Command -ComputerName $fqdn -ErrorAction Stop -ScriptBlock {
                    $excludedFolders = @('Microsoft', 'GoogleSystem')
                    $excludedPrefixes = @(
                        'User_Feed_Synchronization',
                        'MicrosoftEdgeUpdate',
                        'OneDrive Reporting Task',
                        'Optimize Start Menu Cache Files',
                        'OneDrive Startup Task',
                        'SensorFramework-LogonTask'
                    )

                    Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
                        # Exclude tasks in well-known vendor/system folders
                        $folder = $_.TaskPath.Trim('\').Split('\')[0]
                        if ($excludedFolders -contains $folder) { return $false }

                        # Exclude tasks whose name starts with a known prefix
                        $taskName = $_.TaskName
                        $nameMatches = $excludedPrefixes | Where-Object { $taskName -like "$_*" }
                        if ($nameMatches) { return $false }

                        return $true
                    } | ForEach-Object {
                        [PSCustomObject]@{
                            TaskPath  = $_.TaskPath
                            TaskName  = $_.TaskName
                            UserId    = if ($_.Principal) { $_.Principal.UserId }             else { $null }
                            LogonType = if ($_.Principal) { $_.Principal.LogonType.ToString() } else { $null }
                            State     = $_.State.ToString()
                        }
                    }
                }

                $taskRows = $rawTasks | Where-Object {
                    $uid = [string]$_.UserId
                    if ([string]::IsNullOrWhiteSpace($uid))              { return $false }
                    $uid = $uid.Trim()
                    if ($uid -like 'NT AUTHORITY\*')                     { return $false }
                    if ($uid -like 'NT SERVICE\*')                       { return $false }
                    if ($uid -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) { return $false }
                    if ($StandardAccounts -contains $uid.ToLower())      { return $false }
                    return $true
                } | Select-Object `
                    @{N='Server';      E={ $server.Name }},
                    @{N='FQDN';        E={ $fqdn }},
                    @{N='Type';        E={ 'ScheduledTask' }},
                    @{N='Name';        E={ ("$($_.TaskPath)$($_.TaskName)").TrimStart('\') }},
                    @{N='DisplayName'; E={ $_.TaskName }},
                    @{N='StartName';   E={ $_.UserId }},
                    @{N='State';       E={ $_.State }},
                    @{N='StartMode';   E={ $_.LogonType }},
                    @{N='PathName';    E={ '' }}

                Write-Log "$fqdn — Scheduled tasks with non-standard accounts: $($taskRows.Count)"
            }
            catch {
                Write-Log "$fqdn — Could not query scheduled tasks: $_" -Level WARN
            }
        }
        else {
            Write-Log "$fqdn — Scheduled task query skipped (PSRemoting unavailable via DCOM fallback)." -Level WARN
        }

        #--- Export services and tasks to separate CSVs ---
        if ($serviceRows.Count -gt 0) {
            $servicesCsvPath = Join-Path $OutputFolder "$($server.Name)_Services.csv"
            $serviceRows | Export-Csv -Path $servicesCsvPath -NoTypeInformation -Encoding UTF8 -Force  # PS 5.1: writes BOM; use UTF8NoBOM on PS 7+
            $script:ServiceCsvPaths.Add($servicesCsvPath)
            Write-Log "$fqdn — Service accounts exported to $servicesCsvPath ($($serviceRows.Count) row(s))."
        }
        else {
            Write-Log "$fqdn — No non-standard service accounts found."
        }

        if ($taskRows.Count -gt 0) {
            $tasksCsvPath = Join-Path $OutputFolder "$($server.Name)_ScheduledTasks.csv"
            $taskRows | Export-Csv -Path $tasksCsvPath -NoTypeInformation -Encoding UTF8 -Force        # PS 5.1: writes BOM; use UTF8NoBOM on PS 7+
            $script:TaskCsvPaths.Add($tasksCsvPath)
            Write-Log "$fqdn — Scheduled task accounts exported to $tasksCsvPath ($($taskRows.Count) row(s))."
        }
        else {
            Write-Log "$fqdn — No non-standard scheduled task accounts found."
        }
    }
    catch {
        Write-Log "$fqdn — Error querying services: $_" -Level ERROR
        $script:FailedServers.Add($server.Name)
        $totalErrors++
    }
    finally {
        if ($cimSession) {
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
    }
}
Write-Progress -Activity 'Scanning servers' -Completed

#endregion

#region --- Combine CSVs into Excel workbook ---

Write-Log "Combining per-server CSV files into Excel workbook..."
# Use the paths tracked during this run — prevents picking up CSVs from previous runs
$serviceCsvFiles = @($script:ServiceCsvPaths)
$taskCsvFiles    = @($script:TaskCsvPaths)

if ($serviceCsvFiles.Count -eq 0 -and $taskCsvFiles.Count -eq 0) {
    Write-Log "No CSV files produced in this run — skipping Excel export." -Level WARN
}
else {
    try {
        # Remove existing workbook so we start fresh
        if (Test-Path $script:ExcelFile) {
            Remove-Item $script:ExcelFile -Force
        }

        if ($serviceCsvFiles.Count -gt 0) {
            $serviceData = foreach ($csv in ($serviceCsvFiles | Sort-Object)) {
                Import-Csv -Path $csv
            }
            $serviceData | Export-Excel `
                -Path          $script:ExcelFile `
                -WorksheetName 'Services' `
                -AutoSize      `
                -AutoFilter    `
                -FreezeTopRow
            Write-Log "Added 'Services' sheet ($($serviceData.Count) row(s) from $($serviceCsvFiles.Count) server(s))."
        }

        if ($taskCsvFiles.Count -gt 0) {
            $taskData = foreach ($csv in ($taskCsvFiles | Sort-Object)) {
                Import-Csv -Path $csv
            }
            $taskData | Export-Excel `
                -Path          $script:ExcelFile `
                -WorksheetName 'ScheduledTasks' `
                -Append        `
                -AutoSize      `
                -AutoFilter    `
                -FreezeTopRow
            Write-Log "Added 'ScheduledTasks' sheet ($($taskData.Count) row(s) from $($taskCsvFiles.Count) server(s))."
        }

        Write-Log "Excel workbook saved to $script:ExcelFile"
    }
    catch {
        Write-Log "Failed to create Excel workbook: $_" -Level ERROR
    }
}

#endregion

#region --- Summary ---

# Export failed servers list for retry
if ($script:FailedServers.Count -gt 0) {
    $header = "# Failed servers — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') — re-run with: -ComputerList '$script:FailedFile'"
    @($header) + $script:FailedServers.ToArray() |
        Out-File -FilePath $script:FailedFile -Encoding utf8 -Force
    Write-Log "Failed server list exported to $script:FailedFile ($($script:FailedServers.Count) server(s))."
}

Write-Log "---------------------------------------------------"
Write-Log "Scan complete."
Write-Log "  Servers in scope    : $(@($servers).Count)"
Write-Log "  Online / scanned    : $totalOnline"
Write-Log "  Offline / skipped   : $totalOffline"
Write-Log "  Errors during scan  : $totalErrors"
Write-Log "  Failed servers file : $(if ($script:FailedServers.Count -gt 0) { $script:FailedFile } else { 'N/A (no failures)' })"
Write-Log "  Excel workbook      : $(if (Test-Path $script:ExcelFile) { $script:ExcelFile } else { 'N/A (no data)' })"
Write-Log "  Output folder       : $OutputFolder"
Write-Log "  Log file            : $script:LogFile"
Write-Log "---------------------------------------------------"

#endregion
