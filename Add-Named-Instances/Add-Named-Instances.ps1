#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Assigns SQL Server named instance service accounts to the required local security groups.
.DESCRIPTION
    Discovers all SQL Server named instance services (SQL Server engine, SQL Agent, SSIS)
    and ensures their NT Service accounts are members of the correct local security groups
    as required by CIS hardening guidelines:

    SQL Server engine and SQL Agent accounts are added to:
        - AdjustMemoryQuotasForAProcess
        - BypassTraverseChecking
        - ReplaceProcessLevelToken
        - LogonAsAService

    SSIS (MsDtsServer) accounts are added to:
        - BypassTraverseChecking
        - ImpersonateClientAfterAuthentication

    Additionally, all Windows services running as NT Service virtual accounts
    (discovered via Win32_Service) are added to LogonAsAService, as this right
    is a prerequisite for any virtual service account to start and run.

    Any of the five groups that do not yet exist are created automatically.
    Accounts already present in a group are skipped. Supports -WhatIf and -Confirm.

    Every run writes a timestamped log file to the directory specified by -LogDirectory
    (default: C:\Temp). All actions — including successful additions, skipped accounts,
    and errors — are recorded there. When -WhatIf is specified the log is still written
    so the proposed changes are auditable. Log files older than 30 days are removed
    automatically at the end of each run.

    After group memberships are updated, the effective local security policy is verified
    via secedit to confirm each local group is referenced in its corresponding Windows
    User Rights Assignment privilege line. A WARN is logged for any gap, because group
    membership alone has no security effect without the LSA right being granted.
.PARAMETER LogDirectory
    Directory where timestamped log files are written. Created automatically if it does
    not exist. Defaults to C:\Temp.
.OUTPUTS
    PSCustomObject with properties Action (Added/WouldAdd/AlreadyMember/Failed), Account,
    and Group for every account/group combination that was evaluated.
    Log file: <LogDirectory>\Add-Named-Instances_yyyyMMdd_HHmmss.log
.EXAMPLE
    .\Add-Named-Instances.ps1
.EXAMPLE
    .\Add-Named-Instances.ps1 -WhatIf
.EXAMPLE
    .\Add-Named-Instances.ps1 -Verbose
    Shows per-account progress on the console in addition to writing the log file.
.EXAMPLE
    .\Add-Named-Instances.ps1 -LogDirectory 'D:\Logs'
    Writes log files to D:\Logs instead of the default C:\Temp.
.NOTES
    Author  : Marcel Stam
    Date    : 2026-05-21
    Version : 1.9.0
    Requires local administrator privileges.
    Intended for use as part of CIS SQL Server hardening.

    Version history:
        1.0.0 - 2026-05-13 - Initial release.
        1.1.0 - 2026-05-13 - Added explicit param() block; moved group name constants to
                            dedicated Configuration region; added early exit when no instances
                            are found; replaced exit 1 with throw; passed MemberCache
                            explicitly to Add-AccountToGroup; added pipeline result
                            collection and verbose summary.
        1.2.0 - 2026-05-19 - Added Write-Log helper; log file written to C:\Temp with
                            timestamp in filename; all verbose/warning output now also
                            persisted to the log file.
        1.3.0 - 2026-05-19 - Fixed WhatIf bug: Add-AccountToGroup now emits a WouldAdd
                            result object instead of $null when ShouldProcess returns
                            $false, preventing null entries in the results list.
                            Added Write-Log calls for successful Added and AlreadyMember
                            actions. Updated summary to report WouldAdd count.
                            Expanded .DESCRIPTION and .OUTPUTS documentation.
        1.4.0 - 2026-05-19 - Added Remove-OldLogs helper; log files older than 30 days
                            are deleted from C:\Temp at the end of each run.
        1.5.0 - 2026-05-19 - Added [Parameter()] and [ValidateNotNullOrEmpty()] to all
                            function parameters per workspace coding guidelines.
                            Removed Write-Error from Write-Log ERROR level: it created a
                            duplicate error record in $Error alongside the throw that always
                            follows; the log file entry is sufficient.
                            Fixed early-return path: log cleanup now runs before the return
                            when no instances are discovered, so old log files are still
                            pruned on those runs.
        1.6.0 - 2026-05-19 - GPO startup script compatibility: replaced em dash characters
                            in Write-Log string literals with ' - ' to prevent a PS 5.1
                            cp1252 parse error (UTF-8 E2 80 94 decodes to curly-quote in
                            cp1252, terminating the string unexpectedly). Added
                            $ErrorActionPreference = 'Stop' so unexpected errors outside
                            try/catch are caught rather than silently continuing in a
                            non-interactive SYSTEM session.
        1.9.0 - 2026-05-21 - Removed unused State property from $ServiceAccounts query;
                            switched to Select-Object -ExpandProperty StartName for a
                            cleaner string list (improvement 6). Wrapped Get-Service and
                            Get-CimInstance calls in try/catch for friendly error logging
                            (improvement 7). Added #region Verify LSA user-right
                            assignments: exports effective security policy via secedit,
                            resolves each local group SID, and logs WARN for any group
                            not referenced in its corresponding User Rights Assignment
                            privilege line (improvement 8).
        1.8.0 - 2026-05-21 - Fixed stale MemberCache bug: Add-AccountToGroup now updates
                            the in-memory cache after a successful addition, preventing
                            spurious Failed results when the same account is processed by
                            multiple loops (issue 1). Deduplicated $ServiceAccounts by
                            excluding named-instance accounts already handled in earlier
                            loops, eliminating the LogonAsAService overlap (issue 2).
                            Added $env:COMPUTERNAME to the startup log entry so logs are
                            machine-identifiable when collected centrally (improvement 3).
                            Made log directory configurable via -LogDirectory parameter
                            with default C:\Temp (improvement 4). Added exit 1 when
                            Failed results exist for GPO/Task Scheduler failure
                            detection (improvement 5).
        1.7.0 - 2026-05-20 - Compliance review: updated .DESCRIPTION to document
                            LogonAsAService as a fifth target group for SQL engine/agent
                            accounts and NT Service virtual accounts; corrected "four
                            groups" to "five groups"; fixed misleading inline comment on
                            $LogonAsServiceGroup; added missing #endregion for
                            #region Add members to groups.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    # Directory where timestamped log files are written. Created automatically if it does not exist.
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogDirectory = 'C:\Temp'
)

#region Functions
# Helper: writes a timestamped entry to the log file and the appropriate PowerShell stream.
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line      = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Verbose "ERROR: $Message" }  # throw follows; Write-Error would duplicate $Error
        default { Write-Verbose $Message }
    }
}

# Helper: removes log files for this script that are older than the specified retention period.
function Remove-OldLogs {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$LogDirectory,
        [Parameter()]
        [ValidateRange(1, 3650)]
        [int]$RetentionDays = 30
    )
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $old = Get-ChildItem -LiteralPath $LogDirectory -Filter 'Add-Named-Instances_*.log' -File -ErrorAction SilentlyContinue |
           Where-Object { $_.LastWriteTime -lt $cutoff }
    foreach ($file in $old) {
        if ($PSCmdlet.ShouldProcess($file.FullName, 'Remove log file older than 30 days')) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-Log "Removed old log file '$($file.Name)'."
            }
            catch {
                Write-Log -Level WARN "Could not remove old log file '$($file.Name)': $_"
            }
        }
    }
    if ($old.Count -gt 0) {
        Write-Log "Log cleanup complete - $($old.Count) file(s) older than $RetentionDays days processed."
    }
}

# Helper: adds one account to one group and emits a structured result object.
function Add-AccountToGroup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Account,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Group,
        [Parameter(Mandatory)]
        [hashtable]$MemberCache
    )
    if ($MemberCache[$Group].Name -notcontains $Account) {
        # ShouldProcess enables -WhatIf and -Confirm support without extra code.
        if ($PSCmdlet.ShouldProcess($Account, "Add to local group '$Group'")) {
            try {
                Add-LocalGroupMember -Group $Group -Member $Account -ErrorAction Stop
                # Update the in-memory cache immediately so subsequent lookups for this
                # account/group see the correct state and avoid duplicate-member errors.
                $MemberCache[$Group] = @($MemberCache[$Group]) + [PSCustomObject]@{ Name = $Account }
                Write-Log "Added '$Account' to '$Group'."
                [PSCustomObject]@{ Action = 'Added'; Account = $Account; Group = $Group }
            }
            catch {
                Write-Log -Level WARN "Failed to add '$Account' to '$Group': $_"
                [PSCustomObject]@{ Action = 'Failed'; Account = $Account; Group = $Group }
            }
        }
        else {
            # -WhatIf path: emit a result so the caller's list stays consistent (no $null entries).
            Write-Log "WhatIf: would add '$Account' to '$Group'."
            [PSCustomObject]@{ Action = 'WouldAdd'; Account = $Account; Group = $Group }
        }
    }
    else {
        # Account already present; log it and emit an object so the caller gets a complete picture.
        Write-Log "Already member: '$Account' in '$Group'."
        [PSCustomObject]@{ Action = 'AlreadyMember'; Account = $Account; Group = $Group }
    }
}
#endregion

#region Configuration
# Ensure unexpected errors outside try/catch blocks are terminating in a non-interactive session.
$ErrorActionPreference = 'Stop'

# Log file - written to $LogDirectory with a timestamp so each run produces a unique file.
$logDir = $LogDirectory
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$script:LogFile = Join-Path $logDir ("Add-Named-Instances_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
Write-Log "Script started on '$env:COMPUTERNAME'. Log: $($script:LogFile)"

# Target local security groups as defined by CIS hardening requirements for SQL Server.
$AdjustMemoryGroup             = 'AdjustMemoryQuotasForAProcess'
$BypassTraverseCheckingGroup   = 'BypassTraverseChecking'
$ImpersonateClientGroup        = 'ImpersonateClientAfterAuthentication'
$ReplaceProcessLevelTokenGroup = 'ReplaceProcessLevelToken'
$LogonAsServiceGroup           = 'LogonAsAService'  # Used for SQL engine/agent named instances and for all NT Service virtual accounts on the machine.
#endregion

#region Discover SQL named instance service accounts
# Use a generic List to avoid array-rebuild overhead when adding items in a loop.
[System.Collections.Generic.List[string]]$SQLInstances = [System.Collections.Generic.List[string]]::new()
[System.Collections.Generic.List[string]]$SQLAgentInstances = [System.Collections.Generic.List[string]]::new()
[System.Collections.Generic.List[string]]$DTSInstances = [System.Collections.Generic.List[string]]::new()

# Get all services whose display name contains 'SQL'; this is intentionally broad
# so it catches SQL Server, SQL Agent, and SSIS without needing the instance name.
try {
    $services = Get-Service -DisplayName '*SQL*' -ErrorAction Stop
}
catch {
    Write-Log -Level ERROR "Get-Service failed while enumerating SQL services: $_"
    throw
}

# Collect all virtual NT Service accounts in one WMI call; expand to a plain string list
# so the deduplication filter and the member-add loop can reference $account directly.
try {
    $ServiceAccounts = Get-CimInstance -Class Win32_Service -Filter "StartName LIKE 'NT Service%'" -ErrorAction Stop |
                       Select-Object -ExpandProperty StartName
}
catch {
    Write-Log -Level ERROR "Get-CimInstance failed while enumerating NT Service accounts: $_"
    throw
}

foreach ($service in $services) {
    # Include SQL Server engine instances, SQL Agent instances, and SSIS (MsDtsServer*).
    # Named instances appear as 'SQL Server (INSTANCENAME)' / 'SQL Server Agent (INSTANCENAME)'.
    if ($service.DisplayName -like 'SQL Server (*)') {
        $SQLInstances.Add("NT Service\$($service.Name)")
    }
    elseif ($service.DisplayName -like 'SQL Server Agent (*)') {
        $SQLAgentInstances.Add("NT Service\$($service.Name)")
    }
    elseif ($service.Name -like 'MsDtsServer*') {
        $DTSInstances.Add("NT Service\$($service.Name)")
    }
}

# Remove well-known built-in system accounts that cannot (and should not) be added
# to a local group as if they were named-instance service accounts.
$skipAccounts = @('LocalSystem', 'NT AUTHORITY\LocalService', 'NT AUTHORITY\NetworkService')
$SQLInstances = $SQLInstances | Where-Object { $_ -and $_ -notin $skipAccounts }
$SQLAgentInstances = $SQLAgentInstances | Where-Object { $_ -and $_ -notin $skipAccounts }
$DTSInstances = $DTSInstances | Where-Object { $_ -and $_ -notin $skipAccounts }

# Deduplicate $ServiceAccounts: remove any account already present in the named-instance
# lists. Without this, the same NT Service account would be processed twice for
# LogonAsAService — once in the SQL engine/agent loop and again in the ServiceAccounts
# loop. The MemberCache is read before any additions, so the second attempt would call
# Add-LocalGroupMember on a group it just joined, causing a 'member already exists'
# exception that is caught and emitted as a spurious Failed result.
$namedInstanceAccounts = @(@($SQLInstances) + @($SQLAgentInstances) + @($DTSInstances))
$ServiceAccounts = $ServiceAccounts | Where-Object { $namedInstanceAccounts -inotcontains $_ }

Write-Log "Discovered SQL service accounts - Engine: $($SQLInstances.Count), Agent: $($SQLAgentInstances.Count), SSIS: $($DTSInstances.Count), Other: $($ServiceAccounts.Count)"

# Exit early if no named instances were found — nothing to configure.
if ($SQLInstances.Count -eq 0 -and $SQLAgentInstances.Count -eq 0 -and $DTSInstances.Count -eq 0 -and $ServiceAccounts.Count -eq 0) {
    Write-Log -Level WARN 'No SQL Server named instances discovered. Nothing to do.'
    Remove-OldLogs -LogDirectory $logDir -RetentionDays 30
    return
}
#endregion

#region Ensure groups exist
# Collect all target groups and create any that are missing, so the script is idempotent
# and can run on freshly built servers without manual pre-requisites.
$allGroups = @($AdjustMemoryGroup, $BypassTraverseCheckingGroup, $ImpersonateClientGroup, $ReplaceProcessLevelTokenGroup, $LogonAsServiceGroup)
foreach ($group in $allGroups) {
    if (-not (Get-LocalGroup -Name $group -ErrorAction SilentlyContinue)) {
        if ($PSCmdlet.ShouldProcess($group, 'Create local group')) {
            try {
                New-LocalGroup -Name $group -ErrorAction Stop
                Write-Log "Created local group '$group'."
            }
            catch {
                Write-Log -Level ERROR "Failed to create group '$group': $_"
                throw "Failed to create group '$group': $_"
            }
        }
    }
}
#endregion

#region Add members to groups
# Read current membership for every target group once up front to avoid repeated lookups.
$groupMembers = @{}
foreach ($group in $allGroups) {
    try {
        $groupMembers[$group] = Get-LocalGroupMember -Group $group -ErrorAction Stop
    }
    catch {
        Write-Log -Level ERROR "Failed to read members of group '$group': $_"
        throw "Failed to read members of group '$group': $_"
    }
}

# Collect all results so a summary can be emitted and the caller can pipeline/export them.
$results = [System.Collections.Generic.List[object]]::new()

# SQL Server engine and SQL Agent instances require three user-rights assignment groups:
#   AdjustMemoryQuotasForAProcess  — needed to manage memory for SQL processes
#   BypassTraverseChecking         — needed to traverse file paths
#   ReplaceProcessLevelToken       — needed to run processes under the service account
foreach ($account in (@($SQLInstances) + @($SQLAgentInstances))) {
    $results.Add((Add-AccountToGroup -Account $account -Group $AdjustMemoryGroup -MemberCache $groupMembers))
    $results.Add((Add-AccountToGroup -Account $account -Group $BypassTraverseCheckingGroup -MemberCache $groupMembers))
    $results.Add((Add-AccountToGroup -Account $account -Group $ReplaceProcessLevelTokenGroup -MemberCache $groupMembers))
    $results.Add((Add-AccountToGroup -Account $account -Group $LogonAsServiceGroup -MemberCache $groupMembers))
}

# SSIS (DTS) instances require two user-rights assignment groups:
#   BypassTraverseChecking              — needed to traverse file paths
#   ImpersonateClientAfterAuthentication — needed to execute packages under the caller's identity
foreach ($account in $DTSInstances) {
    $results.Add((Add-AccountToGroup -Account $account -Group $BypassTraverseCheckingGroup -MemberCache $groupMembers))
    $results.Add((Add-AccountToGroup -Account $account -Group $ImpersonateClientGroup      -MemberCache $groupMembers))
}

# Add remaining NT Service virtual accounts (named-instance accounts already handled above
# are excluded by the deduplication step in the Discover region) to LogonAsAService.
foreach ($account in $ServiceAccounts) {
    $results.Add((Add-AccountToGroup -Account $account -Group $LogonAsServiceGroup -MemberCache $groupMembers))
}

# Emit all results to the pipeline for the caller to capture, filter, or export.
$results

# Write a concise summary to both the log file and the verbose stream.
Write-Log ("Summary - Added: {0}, WouldAdd: {1}, AlreadyMember: {2}, Failed: {3}" -f
    ($results | Where-Object Action -eq 'Added').Count,
    ($results | Where-Object Action -eq 'WouldAdd').Count,
    ($results | Where-Object Action -eq 'AlreadyMember').Count,
    ($results | Where-Object Action -eq 'Failed').Count)
Write-Log 'Script completed.'

#region Log cleanup
# Remove log files for this script that are older than 30 days to prevent unbounded growth.
Remove-OldLogs -LogDirectory $logDir -RetentionDays 30
#endregion
#endregion

#region Verify LSA user-right assignments
# The local groups this script populates grant Windows User Rights only if the local
# security policy (or an applied GPO) references each group in its User Rights Assignment.
# Export the effective policy via secedit and warn on any gap to surface misconfigurations
# where group memberships are correct but the rights are not actually delegated.
$privilegeMap = [ordered]@{
    $AdjustMemoryGroup             = 'SeIncreaseQuotaPrivilege'
    $BypassTraverseCheckingGroup   = 'SeChangeNotifyPrivilege'
    $ImpersonateClientGroup        = 'SeImpersonatePrivilege'
    $ReplaceProcessLevelTokenGroup = 'SeAssignPrimaryTokenPrivilege'
    $LogonAsServiceGroup           = 'SeServiceLogonRight'
}

$tempCfg = Join-Path $env:TEMP ("secedit_{0}.cfg" -f [System.IO.Path]::GetRandomFileName())
try {
    $null = & secedit /export /cfg $tempCfg /quiet 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Log -Level WARN "secedit /export exited with code $LASTEXITCODE - LSA verification skipped."
    }
    else {
        # secedit writes UTF-16 LE; -Encoding Unicode decodes it correctly on PS 5.1 and 7+.
        $policyLines = Get-Content -LiteralPath $tempCfg -Encoding Unicode -ErrorAction Stop

        $lsaOk   = 0
        $lsaGaps = 0
        foreach ($groupName in $privilegeMap.Keys) {
            $privilege  = $privilegeMap[$groupName]
            $localGroup = Get-LocalGroup -Name $groupName -ErrorAction SilentlyContinue
            if (-not $localGroup) {
                Write-Log -Level WARN "LSA check skipped for '$groupName' - group does not exist (created under -WhatIf and not persisted?)."
                $lsaGaps++
                continue
            }
            $sid      = $localGroup.SID.Value
            $privLine = $policyLines | Where-Object { $_ -match "^\s*$([regex]::Escape($privilege))\s*=" }
            if ($privLine -and $privLine -match "\*$([regex]::Escape($sid))") {
                Write-Log "LSA OK - '$groupName' (SID $sid) is referenced in '$privilege'."
                $lsaOk++
            }
            else {
                Write-Log -Level WARN "LSA GAP - '$groupName' is NOT referenced in '$privilege'. Group membership has no security effect without this User Rights Assignment. Configure it via GPO or Local Security Policy."
                $lsaGaps++
            }
        }
        Write-Log ("LSA verification complete - OK: $lsaOk, Gaps: $lsaGaps.")
    }
}
catch {
    Write-Log -Level WARN "LSA verification failed unexpectedly: $_"
}
finally {
    Remove-Item -LiteralPath $tempCfg -Force -ErrorAction SilentlyContinue
}
#endregion

# Signal failure to the calling process (GPO startup, Task Scheduler) when any action
# failed, so the infrastructure can detect and alert on partial-hardening runs.
if (($results | Where-Object Action -eq 'Failed').Count -gt 0) {
    exit 1
}