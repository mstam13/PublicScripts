<#
.SYNOPSIS
    Reports every configured setting in every GPO in a domain, regardless of where
    (or whether) each GPO is linked.
.DESCRIPTION
    Enumerates all Group Policy Objects in the specified domain, fetches the full
    XML report (Get-GPOReport) for each one, and parses configured settings into a
    single flat list. Covers: Administrative Templates, Security Settings (Account
    Policies, User Rights Assignment, Security Options, Audit Policy, Restricted
    Groups, System Services), Scripts, and Windows Firewall rules.

    Unlike Compare-GPOsByOU.ps1 (which reports settings per OU link), this script is
    scoped purely to GPOs — orphaned/unlinked GPOs are included, and each configured
    setting appears exactly once per GPO regardless of how many containers link it.

    The results are exported to a single-worksheet Excel workbook when the
    ImportExcel module is available, or to a CSV file as a fallback.
    A timestamped log file is always written to a 'Log' sub-folder next to the script.

    Output filenames include the domain name and the run date, e.g.:
      2026-08-20_contoso.com_Get-AllGPOSettings.xlsx
.PARAMETER Domain
    FQDN of the Active Directory domain to query.
    Defaults to the current user's logon domain ($env:USERDNSDOMAIN).
.PARAMETER DomainController
    FQDN or hostname of a specific domain controller to target for all GPO queries.
    When omitted, Windows picks any available DC. Specifying a DC ensures all calls
    hit the same source, preventing inconsistencies caused by replication lag.
.PARAMETER OutputPath
    Directory where the Excel/CSV report is written.
    Defaults to the directory containing this script ($PSScriptRoot).
    The directory is created automatically if it does not exist.
.PARAMETER PassThru
    When specified, returns the collected settings as an array of objects to the
    pipeline in addition to writing the report file.
.OUTPUTS
    <OutputPath>\YYYY-MM-dd_<Domain>_Get-AllGPOSettings.xlsx
      Worksheet 'Settings' — One row per configured setting per GPO.
    (or a single CSV file when ImportExcel is unavailable)
    <PSScriptRoot>\Log\YYYYMMDD_HHmmss_<Domain>_Get-AllGPOSettings.log
.EXAMPLE
    .\Get-AllGPOSettings.ps1

    Runs against the current user's domain and writes output to the script directory.
.EXAMPLE
    .\Get-AllGPOSettings.ps1 -Domain contoso.com -OutputPath C:\Reports

    Runs against contoso.com and writes the report to C:\Reports.
.EXAMPLE
    .\Get-AllGPOSettings.ps1 -Domain contoso.com -DomainController dc01.contoso.com

    Pins all queries to a specific domain controller for consistency.
.EXAMPLE
    $settings = .\Get-AllGPOSettings.ps1 -PassThru
    $settings | Where-Object { $_.ExtensionType -eq 'Administrative Templates' }

    Returns the result data to the pipeline for further filtering.
.NOTES
    Author      : M. Stam
    Date        : 2026-09-09
    Version     : 1.1.0

    Requires    : GroupPolicy module (RSAT-GPMC)
                  ImportExcel module (optional; https://github.com/dfinke/ImportExcel)

    Permissions : Group Policy: Read on all GPOs.

    Each GPO report requires one network round-trip to the DC. On domains with many
    GPOs this may significantly increase runtime. On PowerShell 7+, reports are
    fetched in parallel (up to 8 concurrent requests).

    Settings worksheet columns mirror the 'Settings' worksheet produced by
    Compare-GPOsByOU.ps1, minus the OU-link-specific columns (ContainerName,
    ContainerDN, LinkOrder, LinkEnabled, LinkEnforced) which have no meaning when
    reporting on GPOs independent of their links. GPOId and GPOStatus are added
    for unambiguous identification since a report row is no longer tied to a
    single link.

    A GPO's WmiFilter property can be non-null even when the WMI filter it
    references has been deleted from Active Directory (orphaned gPCWQLFilter
    reference). Get-GpoWmiFilterName wraps the .Name access in try/catch so an
    orphaned filter reference is logged as a warning (WMIFilterName left blank)
    instead of crashing the script with "The property 'Name' cannot be found on
    this object." under Set-StrictMode -Version Latest.

    Version history:
      1.1.0  2026-09-09  M. Stam  Moved Get-XmlInnerText and Get-GpoSetting to shared module
                                   Shared\PublicScripts.psm1 (removed duplication with
                                   Compare-GPOsByOU.ps1). Moved module-check block to shared
                                   Import-RequiredModule.
      1.0.1  2026-08-20  M. Stam  Fixed crash ("The property 'Name' cannot be
                                   found on this object") when a GPO references
                                   a WMI filter that no longer exists in AD;
                                   WMIFilterName resolution now uses try/catch
                                   instead of a bare null-check.
      1.0.0  2026-08-20  M. Stam  Initial release.
#>
#Requires -Modules GroupPolicy

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

Import-Module (Join-Path $script:ScriptRoot '..\Shared\PublicScripts.psm1') -Force
$null = Initialize-ScriptLog -LogDirectory (Join-Path $script:ScriptRoot 'Log') `
    -ScriptName 'Get-AllGPOSettings' -Tag $Domain
#endregion

#region Helper functions
function Get-GpoWmiFilterName {
    <#
    .SYNOPSIS Safely resolves a GPO's WMI filter display name.
    .DESCRIPTION
        $gpo.WmiFilter can be a non-null object even when the WMI filter it references
        has been deleted from Active Directory (an orphaned gPCWQLFilter reference on
        the GPO). In that case, accessing .Name throws a PropertyNotFoundException
        under Set-StrictMode -Version Latest ("The property 'Name' cannot be found on
        this object."). Wrapping the access in try/catch avoids that crash.
    #>
    param ($Gpo)
    if ($null -eq $Gpo.WmiFilter) { return $null }
    try { return $Gpo.WmiFilter.Name }
    catch {
        Write-ScriptLog "Could not resolve WMI filter name for GPO '$($Gpo.DisplayName)' (possibly an orphaned filter reference): $_" -Level 'WARN'
        return $null
    }
}
#endregion

#region Main
# Resolve the server target used for all GPO calls
$adServer  = if ($PSBoundParameters.ContainsKey('DomainController')) { $DomainController } else { $Domain }
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Write-ScriptLog "Starting all-GPO settings report for domain: $Domain (server: $adServer)"

# Import required module
Write-ScriptLog 'Checking required modules...'
Import-RequiredModule -Name 'GroupPolicy'

# Retrieve every GPO in the domain
Write-ScriptLog 'Retrieving all GPOs...'
$allGpos = @(Get-GPO -All -Domain $Domain -Server $adServer)
Write-ScriptLog "Found $($allGpos.Count) GPOs."

$settingsRows = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($allGpos.Count -gt 0) {
    # Pre-compute GPO-level metadata used on every settings row (cheap; no round-trip)
    $gpoMeta = @{}
    foreach ($gpo in $allGpos) {
        $statusStr = $gpo.GpoStatus.ToString()
        $gpoMeta[$gpo.Id.ToString()] = [PSCustomObject]@{
            GpoName                 = $gpo.DisplayName
            GpoId                   = $gpo.Id.ToString('B').ToUpper()
            GpoStatus               = $statusStr
            WMIFilterName           = Get-GpoWmiFilterName -Gpo $gpo
            ComputerSettingsEnabled = $statusStr -notin @('ComputerSettingsDisabled', 'AllSettingsDisabled')
            UserSettingsEnabled     = $statusStr -notin @('UserSettingsDisabled', 'AllSettingsDisabled')
        }
    }

    Write-ScriptLog "Fetching settings reports for $($allGpos.Count) GPOs..."
    $reportCache = @{}

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        # Parallel fetching on PS7+ — pass function bodies via $using:
        Write-ScriptLog '  Using parallel report fetching (PS 7+).'
        $getXmlFnBody = ${function:Get-XmlInnerText}.ToString()
        $getGpoFnBody = ${function:Get-GpoSetting}.ToString()

        $parallelResults = $allGpos | ForEach-Object -Parallel {
            ${function:Get-XmlInnerText} = [scriptblock]::Create($using:getXmlFnBody)
            ${function:Get-GpoSetting}   = [scriptblock]::Create($using:getGpoFnBody)
            $gpo    = $_
            $gpoId  = $gpo.Id.ToString()
            [PSCustomObject]@{
                GpoId    = $gpoId
                Settings = Get-GpoSetting -GpoGuid $gpoId -GpoName $gpo.DisplayName `
                    -DomainFqdn $using:Domain -Server $using:adServer
            }
        } -ThrottleLimit 8

        foreach ($result in $parallelResults) {
            $reportCache[$result.GpoId] = $result.Settings
        }
    }
    else {
        # Sequential fallback for PS 5
        $count = 0
        foreach ($gpo in $allGpos) {
            $count++
            Write-Progress -Activity "Fetching GPO reports ($count / $($allGpos.Count))" `
                -Status $gpo.DisplayName -PercentComplete ([int](($count / $allGpos.Count) * 100))
            $reportCache[$gpo.Id.ToString()] = Get-GpoSetting -GpoGuid $gpo.Id.ToString() `
                -GpoName $gpo.DisplayName -DomainFqdn $Domain -Server $adServer
        }
        Write-Progress -Activity 'Fetching GPO reports' -Completed
    }

    foreach ($gpo in $allGpos) {
        $rawGuid = $gpo.Id.ToString()
        if (-not $reportCache.ContainsKey($rawGuid)) { continue }

        $meta = $gpoMeta[$rawGuid]
        foreach ($setting in $reportCache[$rawGuid]) {
            $settingsRows.Add([PSCustomObject]@{
                GPOName                 = $meta.GpoName
                GPOId                   = $meta.GpoId
                GPOStatus               = $meta.GpoStatus
                WMIFilterName           = $meta.WMIFilterName
                ComputerSettingsEnabled = $meta.ComputerSettingsEnabled
                UserSettingsEnabled     = $meta.UserSettingsEnabled
                Area                    = $setting.Area
                ExtensionType           = $setting.ExtensionType
                Category                = $setting.Category
                SettingName             = $setting.SettingName
                SettingState            = $setting.SettingState
                SettingValue            = $setting.SettingValue
            })
        }
    }
}

# Sort for a consistent, readable spreadsheet
$settingsRows = [System.Collections.Generic.List[PSCustomObject]] @(
    $settingsRows | Sort-Object GPOName, Area, ExtensionType, Category, SettingName
)

Write-ScriptLog "Total settings rows: $($settingsRows.Count)"
Write-ScriptLog "Settings phase complete. [$($stopwatch.Elapsed.ToString('hh\:mm\:ss\.fff'))]"
$stopwatch.Restart()
#endregion

#region Export
$datePrefix = Get-Date -Format 'yyyy-MM-dd'
$baseName   = "${datePrefix}_${Domain}_Get-AllGPOSettings"
$xlsxPath   = Join-Path $OutputPath "${baseName}.xlsx"

$hasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)

if ($hasImportExcel) {
    Write-ScriptLog "Exporting to Excel: $xlsxPath"

    if ($settingsRows.Count -gt 0) {
        $settingsRows | Export-Excel -Path $xlsxPath -WorksheetName 'Settings' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
        Write-ScriptLog "Report written to: $xlsxPath"
    }
    else {
        Write-ScriptLog 'No settings found; no output file written.' -Level 'WARN'
    }
}
else {
    Write-ScriptLog 'ImportExcel not available; falling back to CSV.' -Level 'WARN'

    if ($settingsRows.Count -gt 0) {
        $csvSettings = Join-Path $OutputPath "${baseName}.csv"
        $settingsRows | Export-Csv -Path $csvSettings -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Settings CSV: $csvSettings"
    }
    else {
        Write-ScriptLog 'No settings found; no output file written.' -Level 'WARN'
    }
}
Write-ScriptLog "Export phase complete. [$($stopwatch.Elapsed.ToString('hh\:mm\:ss\.fff'))]"
#endregion

#region PassThru
if ($PassThru) {
    $settingsRows.ToArray()
}
#endregion
