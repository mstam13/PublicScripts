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
    Remove-OldLog         Removes log files in a directory that are older than a configurable
                          retention window.  Supports -WhatIf / -Confirm.
    Import-RequiredModule Imports one or more modules if not already loaded, logging each
                          import via Write-ScriptLog.
    Get-XmlInnerText      Returns the InnerText of a named child XML node, or $null when absent.
    Get-GpoSetting        Parses a GPO XML report (Get-GPOReport) into a flat list of
                          configured settings. Safe to call from ForEach-Object -Parallel
                          runspaces (uses Write-Warning, not Write-ScriptLog).
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

# ---------------------------------------------------------------------------

function Import-RequiredModule {
    <#
    .SYNOPSIS
        Imports one or more modules if they are not already loaded.
    .DESCRIPTION
        Checks each named module with Get-Module and imports it (-ErrorAction Stop) only when
        not already present in the session, logging the import via Write-ScriptLog.
        Write-ScriptLog must be initialised before calling this function.
    .PARAMETER Name
        One or more module names to ensure are imported.
    .EXAMPLE
        Import-RequiredModule -Name 'GroupPolicy', 'ActiveDirectory'
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Name
    )

    foreach ($module in $Name) {
        if (-not (Get-Module -Name $module)) {
            Write-ScriptLog "Importing module: $module"
            Import-Module $module -ErrorAction Stop
        }
    }
}

# ---------------------------------------------------------------------------

function Get-XmlInnerText {
    <#
    .SYNOPSIS
        Returns InnerText of a child XML node, or $null when absent.
    .PARAMETER ParentNode
        The XML node to search for a child element.
    .PARAMETER LocalName
        Local (namespace-agnostic) name of the child element to return the InnerText of.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.Xml.XmlNode] $ParentNode,

        [Parameter(Mandatory)]
        [string] $LocalName
    )
    $child = $ParentNode.SelectSingleNode("*[local-name()='$LocalName']")
    if ($null -ne $child) { return $child.InnerText }
    return $null
}

# ---------------------------------------------------------------------------

function Get-GpoSetting {
    <#
    .SYNOPSIS
        Parses a GPO XML report and returns a flat list of settings.
    .DESCRIPTION
        Fetches the full XML report for a single GPO (Get-GPOReport) and flattens its
        configured settings into a list of objects covering: Administrative Templates,
        Security Settings (Account Policies, User Rights Assignment, Audit Policy,
        Security Options, Restricted Groups, System Services), Scripts, and Windows
        Firewall rules.
    .PARAMETER GpoGuid
        GUID of the GPO to report on.
    .PARAMETER GpoName
        Display name of the GPO, used only for warning messages on failure.
    .PARAMETER DomainFqdn
        FQDN of the domain the GPO belongs to.
    .PARAMETER Server
        Optional domain controller to target for the Get-GPOReport call.
    .NOTES
        Uses Write-Warning (not Write-ScriptLog) so it is safe to call from
        ForEach-Object -Parallel runspaces that cannot write to the shared log file.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $GpoGuid,

        [Parameter(Mandatory)]
        [string] $GpoName,

        [Parameter(Mandatory)]
        [string] $DomainFqdn,

        [Parameter()]
        [string] $Server
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $reportParams = @{
            Guid        = $GpoGuid
            ReportType  = 'Xml'
            Domain      = $DomainFqdn
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrEmpty($Server)) { $reportParams['Server'] = $Server }
        [xml] $xml = Get-GPOReport @reportParams
    }
    catch {
        Write-Warning "Cannot retrieve report for '$GpoName': $_"
        return $results
    }

    foreach ($area in @('Computer', 'User')) {
        $areaNode = $xml.SelectSingleNode("//*[local-name()='$area']")
        if ($null -eq $areaNode) { continue }

        # Administrative Templates
        foreach ($policy in @($areaNode.SelectNodes(".//*[local-name()='Policy']"))) {
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Administrative Templates'
                Category      = Get-XmlInnerText $policy 'Category'
                SettingName   = Get-XmlInnerText $policy 'Name'
                SettingState  = Get-XmlInnerText $policy 'State'
                SettingValue  = $null
            })
        }

        # Security Settings – Account Policies (password, lockout, Kerberos)
        foreach ($account in @($areaNode.SelectNodes(".//*[local-name()='Account']"))) {
            $value = Get-XmlInnerText $account 'SettingNumber'
            if ($null -eq $value) { $value = Get-XmlInnerText $account 'SettingBoolean' }
            if ($null -eq $value) { $value = Get-XmlInnerText $account 'SettingString' }
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Security Settings'
                Category      = 'Account Policies'
                SettingName   = Get-XmlInnerText $account 'Name'
                SettingState  = $null
                SettingValue  = $value
            })
        }

        # Security Settings – User Rights Assignment
        foreach ($ura in @($areaNode.SelectNodes(".//*[local-name()='UserRightsAssignment']"))) {
            $members = @($ura.SelectNodes("*[local-name()='Member']/*[local-name()='Name']")) |
                       ForEach-Object { $_.InnerText }
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Security Settings'
                Category      = 'User Rights Assignment'
                SettingName   = Get-XmlInnerText $ura 'Name'
                SettingState  = $null
                SettingValue  = ($members -join '; ')
            })
        }

        # Security Settings – Audit Policy
        foreach ($audit in @($areaNode.SelectNodes(".//*[local-name()='AuditSetting']"))) {
            $name = Get-XmlInnerText $audit 'SubcategoryName'
            if ($null -eq $name) { $name = Get-XmlInnerText $audit 'Category' }
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Security Settings'
                Category      = 'Audit Policy'
                SettingName   = $name
                SettingState  = $null
                SettingValue  = Get-XmlInnerText $audit 'SettingValue'
            })
        }

        # Security Settings – Security Options (registry-based policy settings)
        foreach ($secOpt in @($areaNode.SelectNodes(".//*[local-name()='SecurityOptions']"))) {
            $displayNode = $secOpt.SelectSingleNode(
                "*[local-name()='Display']/*[local-name()='Name']")
            $displayName = if ($null -ne $displayNode) { $displayNode.InnerText } `
                           else { Get-XmlInnerText $secOpt 'KeyName' }
            $value = Get-XmlInnerText $secOpt 'SettingNumber'
            if ($null -eq $value) { $value = Get-XmlInnerText $secOpt 'SettingString' }
            if ($null -eq $value) { $value = Get-XmlInnerText $secOpt 'SettingBoolean' }
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Security Settings'
                Category      = 'Security Options'
                SettingName   = $displayName
                SettingState  = $null
                SettingValue  = $value
            })
        }

        # Security Settings – Restricted Groups
        foreach ($rg in @($areaNode.SelectNodes(".//*[local-name()='RestrictedGroup']"))) {
            $members = @($rg.SelectNodes(".//*[local-name()='Member']/*[local-name()='Name']")) |
                       ForEach-Object { $_.InnerText }
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Security Settings'
                Category      = 'Restricted Groups'
                SettingName   = Get-XmlInnerText $rg 'GroupName'
                SettingState  = $null
                SettingValue  = ($members -join '; ')
            })
        }

        # Security Settings – System Services
        foreach ($svc in @($areaNode.SelectNodes(".//*[local-name()='NTService']"))) {
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Security Settings'
                Category      = 'System Services'
                SettingName   = Get-XmlInnerText $svc 'ServiceName'
                SettingState  = Get-XmlInnerText $svc 'StartupMode'
                SettingValue  = $null
            })
        }

        # Scripts (Startup / Shutdown / Logon / Logoff)
        foreach ($scriptItem in @($areaNode.SelectNodes(".//*[local-name()='Script']"))) {
            $results.Add([PSCustomObject]@{
                Area          = $area
                ExtensionType = 'Scripts'
                Category      = Get-XmlInnerText $scriptItem 'Type'
                SettingName   = Get-XmlInnerText $scriptItem 'CmdLine'
                SettingState  = 'Configured'
                SettingValue  = Get-XmlInnerText $scriptItem 'Parameters'
            })
        }

        # Windows Firewall Rules (Windows Defender Firewall with Advanced Security)
        foreach ($fwSection in @($areaNode.SelectNodes(".//*[local-name()='FirewallRules']"))) {
            foreach ($rule in @($fwSection.SelectNodes("*[local-name()='Rule']"))) {
                $results.Add([PSCustomObject]@{
                    Area          = $area
                    ExtensionType = 'Windows Firewall'
                    Category      = Get-XmlInnerText $rule 'Profile'
                    SettingName   = Get-XmlInnerText $rule 'Name'
                    SettingState  = Get-XmlInnerText $rule 'Active'
                    SettingValue  = Get-XmlInnerText $rule 'Action'
                })
            }
        }
    }

    return $results
}
