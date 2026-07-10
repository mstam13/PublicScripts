<#
.SYNOPSIS
    Reports all GPOs linked per OU including WMI filter details for side-by-side comparison.
.DESCRIPTION
    Enumerates every Organizational Unit and the domain root in the specified domain,
    collects each GPO link together with its properties (link order, enabled, enforced),
    and resolves each GPO's WMI filter name and WQL query from Active Directory.

    The results are exported to a single Excel workbook (up to six worksheets) when the
    ImportExcel module is available, or to individual CSV files as a fallback.
    A timestamped log file is always written to a 'Log' sub-folder next to the script.

    Output filenames include the domain name and the run date, e.g.:
      2026-07-10_contoso.com_Compare-GPOsByOU.xlsx
.PARAMETER Domain
    FQDN of the Active Directory domain to query.
    Defaults to the current user's logon domain ($env:USERDNSDOMAIN).
.PARAMETER DomainController
    FQDN or hostname of a specific domain controller to target for all AD and GP queries.
    When omitted, Windows picks any available DC. Specifying a DC ensures all calls hit
    the same source, preventing inconsistencies caused by replication lag.
.PARAMETER OutputPath
    Directory where the Excel/CSV report is written.
    Defaults to the directory containing this script ($PSScriptRoot).
    The directory is created automatically if it does not exist.
.PARAMETER SearchBase
    Distinguished Name of the OU to use as the starting point for scanning.
    Only that OU and all OUs beneath it are included; the domain root is excluded.
    Omit this parameter to scan the entire domain (default behaviour).
    Example: 'OU=Offices,DC=contoso,DC=com'
.PARAMETER IncludeAll
    When specified, containers with no linked GPOs are also included in the Summary
    worksheet. By default only containers with at least one GPO link appear.
.PARAMETER CompareSettings
    When specified, fetches the full XML report (Get-GPOReport) for every unique
    linked GPO and parses configured settings into two additional worksheets:
      'Settings'  — One row per configured setting per GPO per container.
                    Covers: Administrative Templates, Security Settings (Account
                    Policies, User Rights Assignment, Security Options, Audit Policy,
                    Restricted Groups, System Services), Scripts, and Windows Firewall
                    rules.
      'Conflicts' — Subset of Settings rows where the same setting is defined in
                    two or more GPOs that are both link-enabled and have the relevant
                    area (Computer/User) enabled on the GPO itself.
    Note: each GPO report requires one network round-trip to the DC. On large
    domains with many linked GPOs this may significantly increase runtime.
    On PowerShell 7+, reports are fetched in parallel (up to 8 concurrent requests).
.PARAMETER PassThru
    When specified, returns the collected data as a hashtable to the pipeline in
    addition to writing the report file. Keys: GPOsByOU, Summary, Orphaned,
    Settings (if -CompareSettings), Conflicts (if -CompareSettings).
.OUTPUTS
    <OutputPath>\YYYY-MM-dd_<Domain>_Compare-GPOsByOU.xlsx
      Worksheet 'GPOsByOU'  — One row per GPO link per container.
      Worksheet 'Summary'   — One row per container with its linked GPO count.
      Worksheet 'Orphaned'  — GPOs that exist in the domain but are not linked anywhere.
      Worksheet 'Settings'  — One row per setting (-CompareSettings only).
      Worksheet 'Conflicts' — Conflicting settings (-CompareSettings only).
    (or up to five CSV files when ImportExcel is unavailable)
    <PSScriptRoot>\Log\YYYYMMDD_HHmmss_<Domain>_Compare-GPOsByOU.log
.EXAMPLE
    .\Compare-GPOsByOU.ps1

    Runs against the current user's domain and writes output to the script directory.
.EXAMPLE
    .\Compare-GPOsByOU.ps1 -Domain contoso.com -OutputPath C:\Reports

    Runs against contoso.com and writes the report to C:\Reports.
.EXAMPLE
    .\Compare-GPOsByOU.ps1 -Domain contoso.com -DomainController dc01.contoso.com

    Pins all queries to a specific domain controller for consistency.
.EXAMPLE
    .\Compare-GPOsByOU.ps1 -SearchBase 'OU=Offices,DC=contoso,DC=com'

    Scans only the Offices OU and all OUs beneath it.
.EXAMPLE
    .\Compare-GPOsByOU.ps1 -IncludeAll

    Includes containers with no linked GPOs in the Summary worksheet.
.EXAMPLE
    .\Compare-GPOsByOU.ps1 -SearchBase 'OU=IT,DC=contoso,DC=com' -CompareSettings

    Scans the IT OU subtree and compares configured GPO settings,
    reporting conflicts where the same setting is set by multiple GPOs.
.EXAMPLE
    $data = .\Compare-GPOsByOU.ps1 -PassThru
    $data.Conflicts | Where-Object { $_.ExtensionType -eq 'Administrative Templates' }

    Returns the result data to the pipeline for further filtering.
.NOTES
    Author      : M. Stam
    Date        : 2026-07-10
    Version     : 1.3.0

    Requires    : GroupPolicy module (RSAT-GPMC)
                  ActiveDirectory module (RSAT-AD-PowerShell)
                  ImportExcel module (optional; https://github.com/dfinke/ImportExcel)

    Permissions : Domain read access.
                  Group Policy: Read on all GPOs.
                  Active Directory: Read access to
                    CN=SOM,CN=WMIPolicy,CN=System,<DomainDN>
                  for WMI filter query retrieval. If access is denied, WMIFilterQuery
                  is left blank and a warning is logged.

    WMI filter details (name and WQL query) are read from msWMI-Som objects stored in
    CN=SOM,CN=WMIPolicy,CN=System,<DomainDN>. The msWMI-Parm2 attribute contains
    the WMI namespace and WQL query separated by a semicolon, for example:
      root\CIMv2;SELECT * FROM Win32_OperatingSystem WHERE Version LIKE "10.%"

    Version history:
      1.3.0  2026-07-10  M. Stam  Added -DomainController, -PassThru; Write-Progress;
                                   orphaned GPO detection; Security Options, Restricted
                                   Groups, System Services, Windows Firewall rules in
                                   -CompareSettings; improved conflict detection
                                   (respects Computer/User area enable state); parallel
                                   GPO report fetching on PS7+; WMI filter null guard;
                                   fixed double-logging; OutputPath auto-create.
      1.2.0  2026-07-07  M. Stam  Added -CompareSettings: GPO settings extraction
                                   and conflict detection per OU.
      1.1.0  2026-07-07  M. Stam  Added -SearchBase to scope scanning to a subtree.
      1.0.0  2026-07-02  M. Stam  Initial release.
#>
#Requires -Modules GroupPolicy, ActiveDirectory

[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $Domain = $env:USERDNSDOMAIN,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $DomainController,

    [Parameter()]
    [string] $OutputPath = $PSScriptRoot,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $SearchBase,

    [Parameter()]
    [switch] $IncludeAll,

    [Parameter()]
    [switch] $CompareSettings,

    [Parameter()]
    [switch] $PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Logging
# Guard against empty $PSScriptRoot (e.g. when pasted into an interactive console)
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
if ([string]::IsNullOrEmpty($OutputPath)) { $OutputPath = $script:ScriptRoot }

# Ensure OutputPath exists
if (-not (Test-Path -Path $OutputPath -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $OutputPath -Force
}

$LogDir  = Join-Path $script:ScriptRoot 'Log'
if (-not (Test-Path $LogDir)) { $null = New-Item -ItemType Directory -Path $LogDir }
$LogFile = Join-Path $LogDir "$(Get-Date -Format 'yyyyMMdd_HHmmss')_${Domain}_Compare-GPOsByOU.log"

function Write-ScriptLog {
    param ([string] $Message, [string] $Level = 'INFO')
    $entry = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $entry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    if ($Level -eq 'ERROR')    { Write-Error   $Message }
    elseif ($Level -eq 'WARN') { Write-Warning $Message }
    else { Write-Information -MessageData $entry -InformationAction Continue }
}
#endregion

#region Helper functions
function Get-XmlInnerText {
    <#
    .SYNOPSIS Returns InnerText of a child XML node, or $null when absent.
    #>
    param ([System.Xml.XmlNode] $ParentNode, [string] $LocalName)
    $child = $ParentNode.SelectSingleNode("*[local-name()='$LocalName']")
    if ($null -ne $child) { return $child.InnerText }
    return $null
}

function Get-GpoSetting {
    <#
    .SYNOPSIS Parses a GPO XML report and returns a flat list of settings.
    .NOTES Uses Write-Warning (not Write-ScriptLog) so it is safe to call from
           ForEach-Object -Parallel runspaces that cannot write to the shared log file.
    #>
    param (
        [string] $GpoGuid,
        [string] $GpoName,
        [string] $DomainFqdn,
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
#endregion

#region Main
# Resolve the server target used for all AD and GPO calls
$adServer = if ($PSBoundParameters.ContainsKey('DomainController')) { $DomainController } else { $Domain }

Write-ScriptLog "Starting GPO-by-OU comparison for domain: $Domain (server: $adServer)"

# Import required modules
Write-ScriptLog 'Checking required modules...'
foreach ($mod in 'GroupPolicy', 'ActiveDirectory') {
    if (-not (Get-Module -Name $mod)) {
        Write-ScriptLog "Importing module: $mod"
        Import-Module $mod -ErrorAction Stop
    }
}

# Pre-load all GPOs into a lookup hashtable (key = GUID string, lowercase without braces)
Write-ScriptLog 'Retrieving all GPOs...'
$gpoTable = @{}
foreach ($g in Get-GPO -All -Domain $Domain -Server $adServer) {
    $gpoTable[$g.Id.ToString()] = $g
}
Write-ScriptLog "Found $($gpoTable.Count) GPOs."

# Load all WMI filters from AD, keyed by display name; warn on null names or duplicates
Write-ScriptLog 'Loading WMI filters from Active Directory...'
$wmiFilterTable = @{}   # key = msWMI-Name
$domainDN = (Get-ADDomain -Server $adServer).DistinguishedName
try {
    $somPath    = "CN=SOM,CN=WMIPolicy,CN=System,$domainDN"
    $wmiObjects = Get-ADObject -SearchBase $somPath `
        -Filter { objectClass -eq 'msWMI-Som' } `
        -Properties 'msWMI-Name', 'msWMI-Parm1', 'msWMI-Parm2', 'msWMI-ID' `
        -Server $adServer

    foreach ($wf in $wmiObjects) {
        $filterName = $wf.'msWMI-Name'
        if ([string]::IsNullOrEmpty($filterName)) {
            Write-ScriptLog "WMI filter ID '$($wf.'msWMI-ID')' has no name; skipping." -Level 'WARN'
            continue
        }
        if ($wmiFilterTable.ContainsKey($filterName)) {
            Write-ScriptLog "Duplicate WMI filter name '$filterName' (ID: $($wf.'msWMI-ID')); existing entry overwritten." -Level 'WARN'
        }
        $wmiFilterTable[$filterName] = [PSCustomObject]@{
            Name        = $filterName
            Description = $wf.'msWMI-Parm1'
            Query       = $wf.'msWMI-Parm2'
            ID          = $wf.'msWMI-ID'
        }
    }
    Write-ScriptLog "Found $($wmiFilterTable.Count) WMI filters."
}
catch {
    Write-ScriptLog "Could not load WMI filters from '$domainDN': $_" -Level 'WARN'
}

# Build target list
$targets = [System.Collections.Generic.List[hashtable]]::new()

if ($PSBoundParameters.ContainsKey('SearchBase')) {
    # Validate that the SearchBase OU exists
    Write-ScriptLog "Enumerating containers (scoped to: $SearchBase)..."
    try {
        $null = Get-ADOrganizationalUnit -Identity $SearchBase -Server $adServer -ErrorAction Stop
    }
    catch {
        Write-ScriptLog "SearchBase '$SearchBase' not found or inaccessible: $_" -Level 'ERROR'
    }

    # Include the SearchBase OU itself and all descendants (-SearchScope Subtree is the default)
    foreach ($ou in (Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -Server $adServer)) {
        $targets.Add(@{
            Name              = $ou.Name
            DistinguishedName = $ou.DistinguishedName
            GPITarget         = $ou.DistinguishedName
        })
    }
}
else {
    # Full domain scan: domain root + all OUs
    Write-ScriptLog 'Enumerating containers (domain root + all OUs)...'

    # Domain root — Get-GPInheritance requires the FQDN (not the DN) for the domain container
    $targets.Add(@{
        Name              = $Domain
        DistinguishedName = $domainDN
        GPITarget         = $Domain
    })

    foreach ($ou in (Get-ADOrganizationalUnit -Filter * -Properties 'Name' -Server $adServer)) {
        $targets.Add(@{
            Name              = $ou.Name
            DistinguishedName = $ou.DistinguishedName
            GPITarget         = $ou.DistinguishedName
        })
    }
}
Write-ScriptLog "Found $($targets.Count) containers to process."

$detailRows     = [System.Collections.Generic.List[PSCustomObject]]::new()
$summaryRows    = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalTargets   = $targets.Count
$processedCount = 0

foreach ($target in $targets) {
    $processedCount++
    Write-Progress -Activity "Scanning containers ($processedCount / $totalTargets)" `
        -Status $target.DistinguishedName `
        -PercentComplete ([int](($processedCount / $totalTargets) * 100))

    Write-ScriptLog "Processing: $($target.DistinguishedName)"

    $links = @()
    try {
        $inheritance = Get-GPInheritance -Target $target.GPITarget -Domain $Domain -Server $adServer -ErrorAction Stop
        $links = $inheritance.GpoLinks
    }
    catch {
        Write-ScriptLog "Could not get GPO inheritance for '$($target.DistinguishedName)': $_" -Level 'WARN'
    }

    $linkedCount = @($links).Count

    if ($IncludeAll -or $linkedCount -gt 0) {
        $summaryRows.Add([PSCustomObject]@{
            ContainerName  = $target.Name
            ContainerDN    = $target.DistinguishedName
            LinkedGPOCount = $linkedCount
        })
    }

    foreach ($link in $links) {
        $gpoGuidStr = $link.GpoId.ToString()
        $gpo        = if ($gpoTable.ContainsKey($gpoGuidStr)) { $gpoTable[$gpoGuidStr] } else { $null }

        $wmiFilterName  = $null
        $wmiFilterDesc  = $null
        $wmiFilterQuery = $null

        if ($null -ne $gpo -and $null -ne $gpo.WmiFilter) {
            $wmiFilterName = $gpo.WmiFilter.Name
            if (-not [string]::IsNullOrEmpty($wmiFilterName) -and $wmiFilterTable.ContainsKey($wmiFilterName)) {
                $wmiFilterDesc  = $wmiFilterTable[$wmiFilterName].Description
                $wmiFilterQuery = $wmiFilterTable[$wmiFilterName].Query
            }
        }

        $gpoStatus       = if ($null -ne $gpo) { $gpo.GpoStatus.ToString() } else { 'Unknown' }
        $computerEnabled = if ($null -ne $gpo) { $gpo.GpoStatus.ToString() -notin @('ComputerSettingsDisabled', 'AllSettingsDisabled') } else { $null }
        $userEnabled     = if ($null -ne $gpo) { $gpo.GpoStatus.ToString() -notin @('UserSettingsDisabled', 'AllSettingsDisabled') } else { $null }

        $detailRows.Add([PSCustomObject]@{
            ContainerName           = $target.Name
            ContainerDN             = $target.DistinguishedName
            LinkOrder               = $link.Order
            LinkEnabled             = $link.Enabled
            LinkEnforced            = $link.Enforced
            GPOName                 = $link.DisplayName
            GPOId                   = $link.GpoId.ToString('B').ToUpper()
            GPOStatus               = $gpoStatus
            ComputerSettingsEnabled = $computerEnabled
            UserSettingsEnabled     = $userEnabled
            WMIFilterName           = $wmiFilterName
            WMIFilterDescription    = $wmiFilterDesc
            WMIFilterQuery          = $wmiFilterQuery
            GPOCreationTime         = if ($null -ne $gpo) { $gpo.CreationTime } else { $null }
            GPOModificationTime     = if ($null -ne $gpo) { $gpo.ModificationTime } else { $null }
            GPODescription          = if ($null -ne $gpo) { $gpo.Description } else { $null }
        })
    }
}
Write-Progress -Activity 'Scanning containers' -Completed

# Sort for a consistent, readable spreadsheet
$detailRows = [System.Collections.Generic.List[PSCustomObject]] @(
    $detailRows | Sort-Object ContainerDN, LinkOrder
)

Write-ScriptLog "Total GPO links found : $($detailRows.Count)"
Write-ScriptLog "Containers with GPOs  : $(($summaryRows | Where-Object { $_.LinkedGPOCount -gt 0 }).Count)"

# Detect orphaned GPOs (exist in the domain but linked nowhere in the scanned scope)
$linkedGpoGuids = @{}
foreach ($row in $detailRows) { $linkedGpoGuids[$row.GPOId.Trim('{', '}').ToLower()] = $true }

$orphanedRows = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($kvp in $gpoTable.GetEnumerator()) {
    if (-not $linkedGpoGuids.ContainsKey($kvp.Key.ToLower())) {
        $orphanedRows.Add([PSCustomObject]@{
            GPOName             = $kvp.Value.DisplayName
            GPOId               = $kvp.Value.Id.ToString('B').ToUpper()
            GPOStatus           = $kvp.Value.GpoStatus.ToString()
            GPOCreationTime     = $kvp.Value.CreationTime
            GPOModificationTime = $kvp.Value.ModificationTime
            GPODescription      = $kvp.Value.Description
        })
    }
}
Write-ScriptLog "Unlinked (orphaned) GPOs : $($orphanedRows.Count)"
#endregion

#region Settings comparison
$settingsRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$conflictRows = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($CompareSettings -and $detailRows.Count -gt 0) {
    $uniqueGpoIds = $detailRows | Select-Object -ExpandProperty GPOId -Unique
    Write-ScriptLog "Fetching settings reports for $($uniqueGpoIds.Count) unique GPOs..."
    $reportCache = @{}

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        # Parallel fetching on PS7+ — pass function bodies via $using:
        Write-ScriptLog '  Using parallel report fetching (PS 7+).'
        $getXmlFnBody = ${function:Get-XmlInnerText}.ToString()
        $getGpoFnBody = ${function:Get-GpoSetting}.ToString()

        $parallelResults = $uniqueGpoIds | ForEach-Object -Parallel {
            ${function:Get-XmlInnerText} = [scriptblock]::Create($using:getXmlFnBody)
            ${function:Get-GpoSetting}   = [scriptblock]::Create($using:getGpoFnBody)
            $gpoId   = $_
            $rawGuid = $gpoId.Trim('{', '}')
            $gTable  = $using:gpoTable
            $dom     = $using:Domain
            $srv     = $using:adServer
            $gpoName = if ($gTable.ContainsKey($rawGuid)) { $gTable[$rawGuid].DisplayName } else { $gpoId }
            [PSCustomObject]@{
                GpoId    = $gpoId
                Settings = Get-GpoSetting -GpoGuid $rawGuid -GpoName $gpoName -DomainFqdn $dom -Server $srv
            }
        } -ThrottleLimit 8

        foreach ($result in $parallelResults) {
            $reportCache[$result.GpoId] = $result.Settings
        }
    }
    else {
        # Sequential fallback for PS 5
        foreach ($gpoId in $uniqueGpoIds) {
            $rawGuid = $gpoId.Trim('{', '}')
            $gpoName = if ($gpoTable.ContainsKey($rawGuid)) {
                $gpoTable[$rawGuid].DisplayName
            }
            else { $gpoId }
            Write-ScriptLog "  Report: $gpoName"
            $reportCache[$gpoId] = Get-GpoSetting -GpoGuid $rawGuid -GpoName $gpoName -DomainFqdn $Domain -Server $adServer
        }
    }

    foreach ($row in $detailRows) {
        if (-not $reportCache.ContainsKey($row.GPOId)) { continue }

        foreach ($setting in $reportCache[$row.GPOId]) {
            $settingsRows.Add([PSCustomObject]@{
                ContainerName           = $row.ContainerName
                ContainerDN             = $row.ContainerDN
                LinkOrder               = $row.LinkOrder
                LinkEnabled             = $row.LinkEnabled
                LinkEnforced            = $row.LinkEnforced
                GPOName                 = $row.GPOName
                WMIFilterName           = $row.WMIFilterName
                ComputerSettingsEnabled = $row.ComputerSettingsEnabled
                UserSettingsEnabled     = $row.UserSettingsEnabled
                Area                    = $setting.Area
                ExtensionType           = $setting.ExtensionType
                Category                = $setting.Category
                SettingName             = $setting.SettingName
                SettingState            = $setting.SettingState
                SettingValue            = $setting.SettingValue
            })
        }
    }
    Write-ScriptLog "Total settings rows      : $($settingsRows.Count)"

    # Conflicts: same Container + Area + SettingName in >1 enabled GPO
    # Only consider rows where the link is enabled AND the relevant area is active on the GPO
    $conflictGroups = $settingsRows |
        Where-Object {
            $_.LinkEnabled -and (
                ($_.Area -eq 'Computer' -and $_.ComputerSettingsEnabled) -or
                ($_.Area -eq 'User'     -and $_.UserSettingsEnabled)
            )
        } |
        Group-Object -Property ContainerDN, Area, SettingName |
        Where-Object {
            @($_.Group | Select-Object -ExpandProperty GPOName -Unique).Count -gt 1
        }

    foreach ($grp in @($conflictGroups)) {
        foreach ($item in ($grp.Group | Sort-Object LinkOrder)) {
            $conflictRows.Add($item)
        }
    }
    Write-ScriptLog "Conflicting setting rows : $($conflictRows.Count)"
}
#endregion

#region Export
$datePrefix = Get-Date -Format 'yyyy-MM-dd'
$baseName   = "${datePrefix}_${Domain}_Compare-GPOsByOU"
$xlsxPath   = Join-Path $OutputPath "${baseName}.xlsx"

$hasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)

if ($hasImportExcel) {
    Write-ScriptLog "Exporting to Excel: $xlsxPath"

    if ($detailRows.Count -gt 0) {
        $detailRows | Export-Excel -Path $xlsxPath -WorksheetName 'GPOsByOU' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($summaryRows.Count -gt 0) {
        $summaryRows | Export-Excel -Path $xlsxPath -WorksheetName 'Summary' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($orphanedRows.Count -gt 0) {
        $orphanedRows | Export-Excel -Path $xlsxPath -WorksheetName 'Orphaned' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($settingsRows.Count -gt 0) {
        $settingsRows | Export-Excel -Path $xlsxPath -WorksheetName 'Settings' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($conflictRows.Count -gt 0) {
        $conflictRows | Export-Excel -Path $xlsxPath -WorksheetName 'Conflicts' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($detailRows.Count -eq 0 -and $summaryRows.Count -eq 0) {
        Write-ScriptLog 'No data found; no output file written.' -Level 'WARN'
    }
    else {
        Write-ScriptLog "Report written to: $xlsxPath"
    }
}
else {
    Write-ScriptLog 'ImportExcel not available; falling back to CSV.' -Level 'WARN'

    if ($detailRows.Count -gt 0) {
        $csvDetail = Join-Path $OutputPath "${baseName}_GPOsByOU.csv"
        $detailRows | Export-Csv -Path $csvDetail -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Detail CSV       : $csvDetail"
    }

    if ($summaryRows.Count -gt 0) {
        $csvSummary = Join-Path $OutputPath "${baseName}_Summary.csv"
        $summaryRows | Export-Csv -Path $csvSummary -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Summary CSV      : $csvSummary"
    }

    if ($orphanedRows.Count -gt 0) {
        $csvOrphaned = Join-Path $OutputPath "${baseName}_Orphaned.csv"
        $orphanedRows | Export-Csv -Path $csvOrphaned -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Orphaned CSV     : $csvOrphaned"
    }

    if ($settingsRows.Count -gt 0) {
        $csvSettings = Join-Path $OutputPath "${baseName}_Settings.csv"
        $settingsRows | Export-Csv -Path $csvSettings -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Settings CSV     : $csvSettings"
    }

    if ($conflictRows.Count -gt 0) {
        $csvConflicts = Join-Path $OutputPath "${baseName}_Conflicts.csv"
        $conflictRows | Export-Csv -Path $csvConflicts -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Conflicts CSV    : $csvConflicts"
    }

    if ($detailRows.Count -eq 0 -and $summaryRows.Count -eq 0) {
        Write-ScriptLog 'No data found; no output file written.' -Level 'WARN'
    }
}
#endregion

#region PassThru
if ($PassThru) {
    $output = [ordered]@{
        GPOsByOU = $detailRows.ToArray()
        Summary  = $summaryRows.ToArray()
        Orphaned = $orphanedRows.ToArray()
    }
    if ($CompareSettings) {
        $output['Settings']  = $settingsRows.ToArray()
        $output['Conflicts'] = $conflictRows.ToArray()
    }
    [PSCustomObject] $output
}
#endregion
