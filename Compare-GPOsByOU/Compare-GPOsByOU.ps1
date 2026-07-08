<#
.SYNOPSIS
    Reports all GPOs linked per OU including WMI filter details for side-by-side comparison.
.DESCRIPTION
    Enumerates every Organizational Unit and the domain root in the specified domain,
    collects each GPO link together with its properties (link order, enabled, enforced),
    and resolves each GPO's WMI filter name and WQL query from Active Directory.

    The results are exported to a single Excel workbook (two worksheets) when the
    ImportExcel module is available, or to individual CSV files as a fallback.
    A timestamped log file is always written to a 'Log' sub-folder next to the script.

    Output filenames include the domain name and the run date, e.g.:
      2026-07-02_contoso.com_Compare-GPOsByOU.xlsx
.PARAMETER Domain
    FQDN of the Active Directory domain to query.
    Defaults to the current user's logon domain ($env:USERDNSDOMAIN).
.PARAMETER OutputPath
    Directory where the Excel/CSV report is written.
    Defaults to the directory containing this script ($PSScriptRoot).
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
                    Policies, User Rights Assignment, Audit Policy), and Scripts.
      'Conflicts' — Subset of Settings rows where the same setting is defined in
                    two or more enabled GPOs linked to the same container.
    Note: each GPO report requires one network round-trip to the DC. On large
    domains with many linked GPOs this may significantly increase runtime.
.OUTPUTS
    <OutputPath>\YYYY-MM-dd_<Domain>_Compare-GPOsByOU.xlsx
      Worksheet 'GPOsByOU'  — One row per GPO link per container.
      Worksheet 'Summary'   — One row per container with its linked GPO count.
      Worksheet 'Settings'  — One row per setting (-CompareSettings only).
      Worksheet 'Conflicts' — Conflicting settings (-CompareSettings only).
    (or up to four CSV files when ImportExcel is unavailable)
    <PSScriptRoot>\Log\YYYYMMDD_HHmmss_<Domain>_Compare-GPOsByOU.log
.EXAMPLE
    .\Compare-GPOsByOU.ps1

    Runs against the current user's domain and writes output to the script directory.
.EXAMPLE
    .\Compare-GPOsByOU.ps1 -Domain contoso.com -OutputPath C:\Reports

    Runs against contoso.com and writes the report to C:\Reports.
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
.NOTES
    Author      : M. Stam
    Date        : 2026-07-07
    Version     : 1.2.0

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
    [string] $OutputPath = $PSScriptRoot,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $SearchBase,

    [Parameter()]
    [switch] $IncludeAll,

    [Parameter()]
    [switch] $CompareSettings
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Logging
$LogDir = Join-Path $PSScriptRoot 'Log'
if (-not (Test-Path $LogDir)) { $null = New-Item -ItemType Directory -Path $LogDir }
$LogFile = Join-Path $LogDir "$(Get-Date -Format 'yyyyMMdd_HHmmss')_${Domain}_Compare-GPOsByOU.log"

function Write-ScriptLog {
    param ([string] $Message, [string] $Level = 'INFO')
    $entry = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $entry | Tee-Object -FilePath $LogFile -Append | Write-Verbose
    if ($Level -eq 'ERROR') { Write-Error $Message }
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
    #>
    param (
        [string] $GpoGuid,
        [string] $GpoName,
        [string] $DomainFqdn
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        [xml] $xml = Get-GPOReport -Guid $GpoGuid -ReportType Xml -Domain $DomainFqdn -ErrorAction Stop
    }
    catch {
        Write-ScriptLog "Cannot retrieve report for '$GpoName': $_" -Level 'WARN'
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
    }

    return $results
}
#endregion

#region Main
Write-ScriptLog "Starting GPO-by-OU comparison for domain: $Domain"

# Import required modules
Write-ScriptLog 'Checking required modules...'
foreach ($mod in 'GroupPolicy', 'ActiveDirectory') {
    if (-not (Get-Module -Name $mod)) {
        Write-ScriptLog "Importing module: $mod"
        Import-Module $mod -ErrorAction Stop
    }
}

# Pre-load all GPOs into a lookup hashtable (key = GUID string)
Write-ScriptLog 'Retrieving all GPOs...'
$gpoTable = @{}
foreach ($g in Get-GPO -All -Domain $Domain) {
    $gpoTable[$g.Id.ToString()] = $g
}
Write-ScriptLog "Found $($gpoTable.Count) GPOs."

# Load all WMI filters from AD (CN=SOM,CN=WMIPolicy,CN=System,<DomainDN>)
Write-ScriptLog 'Loading WMI filters from Active Directory...'
$wmiFilterTable = @{}   # key = msWMI-Name
$domainDN = (Get-ADDomain -Server $Domain).DistinguishedName
try {
    $somPath = "CN=SOM,CN=WMIPolicy,CN=System,$domainDN"
    $wmiObjects = Get-ADObject -SearchBase $somPath `
        -Filter { objectClass -eq 'msWMI-Som' } `
        -Properties 'msWMI-Name', 'msWMI-Parm1', 'msWMI-Parm2', 'msWMI-ID' `
        -Server $Domain

    foreach ($wf in $wmiObjects) {
        $filterName = $wf.'msWMI-Name'
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
        $null = Get-ADOrganizationalUnit -Identity $SearchBase -Server $Domain -ErrorAction Stop
    }
    catch {
        Write-ScriptLog "SearchBase '$SearchBase' not found or inaccessible: $_" -Level 'ERROR'
    }

    # Include the SearchBase OU itself and all descendants (-SearchScope Subtree is the default)
    foreach ($ou in (Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -Server $Domain)) {
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

    foreach ($ou in (Get-ADOrganizationalUnit -Filter * -Properties 'Name' -Server $Domain)) {
        $targets.Add(@{
            Name              = $ou.Name
            DistinguishedName = $ou.DistinguishedName
            GPITarget         = $ou.DistinguishedName
        })
    }
}
Write-ScriptLog "Found $($targets.Count) containers to process."

$detailRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$summaryRows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($target in $targets) {
    Write-ScriptLog "Processing: $($target.DistinguishedName)"

    $links = @()
    try {
        $inheritance = Get-GPInheritance -Target $target.GPITarget -Domain $Domain -ErrorAction Stop
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
            if ($wmiFilterTable.ContainsKey($wmiFilterName)) {
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

Write-ScriptLog "Total GPO links found : $($detailRows.Count)"
Write-ScriptLog "Containers with GPOs  : $(($summaryRows | Where-Object { $_.LinkedGPOCount -gt 0 }).Count)"
#endregion

#region Settings comparison
$settingsRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$conflictRows = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($CompareSettings -and $detailRows.Count -gt 0) {
    $uniqueGpoIds = $detailRows | Select-Object -ExpandProperty GPOId -Unique
    Write-ScriptLog "Fetching settings reports for $($uniqueGpoIds.Count) unique GPOs..."
    $reportCache = @{}

    foreach ($gpoId in $uniqueGpoIds) {
        $rawGuid = $gpoId.Trim('{', '}')
        $gpoName = if ($gpoTable.ContainsKey($rawGuid)) {
            $gpoTable[$rawGuid].DisplayName
        }
        else { $gpoId }
        Write-ScriptLog "  Report: $gpoName"
        $reportCache[$gpoId] = Get-GpoSetting -GpoGuid $rawGuid -GpoName $gpoName -DomainFqdn $Domain
    }

    foreach ($row in $detailRows) {
        if (-not $reportCache.ContainsKey($row.GPOId)) { continue }

        foreach ($setting in $reportCache[$row.GPOId]) {
            $settingsRows.Add([PSCustomObject]@{
                ContainerName = $row.ContainerName
                ContainerDN   = $row.ContainerDN
                LinkOrder     = $row.LinkOrder
                LinkEnabled   = $row.LinkEnabled
                LinkEnforced  = $row.LinkEnforced
                GPOName       = $row.GPOName
                WMIFilterName = $row.WMIFilterName
                Area          = $setting.Area
                ExtensionType = $setting.ExtensionType
                Category      = $setting.Category
                SettingName   = $setting.SettingName
                SettingState  = $setting.SettingState
                SettingValue  = $setting.SettingValue
            })
        }
    }
    Write-ScriptLog "Total settings rows      : $($settingsRows.Count)"

    # Conflicts: same Container + Area + SettingName in >1 enabled GPO
    $conflictGroups = $settingsRows |
        Where-Object { $_.LinkEnabled } |
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
$csvDetail  = Join-Path $OutputPath "${baseName}_GPOsByOU.csv"
$csvSummary = Join-Path $OutputPath "${baseName}_Summary.csv"

$hasImportExcel = Get-Module -ListAvailable -Name ImportExcel

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
        $detailRows | Export-Csv -Path $csvDetail -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Detail CSV  : $csvDetail"
    }

    if ($summaryRows.Count -gt 0) {
        $summaryRows | Export-Csv -Path $csvSummary -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Summary CSV : $csvSummary"
    }

    if ($settingsRows.Count -gt 0) {
        $csvSettings = Join-Path $OutputPath "${baseName}_Settings.csv"
        $settingsRows | Export-Csv -Path $csvSettings -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Settings CSV : $csvSettings"
    }

    if ($conflictRows.Count -gt 0) {
        $csvConflicts = Join-Path $OutputPath "${baseName}_Conflicts.csv"
        $conflictRows | Export-Csv -Path $csvConflicts -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Conflicts CSV: $csvConflicts"
    }

    if ($detailRows.Count -eq 0 -and $summaryRows.Count -eq 0) {
        Write-ScriptLog 'No data found; no output file written.' -Level 'WARN'
    }
}
#endregion
