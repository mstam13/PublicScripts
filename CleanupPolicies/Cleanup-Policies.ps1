<#
.SYNOPSIS
    Inventories GPOs that are unlinked or have no 'Apply Group Policy' ACE.
.DESCRIPTION
    Queries all Group Policy Objects in the current (or specified) domain and
    reports two categories of potentially unused policies:

      1. Unlinked GPOs  — GPOs with no link to any OU, site, or domain container.
      2. No Apply ACE   — GPOs whose security descriptor contains no Allow ACE for
                          the 'Apply Group Policy' extended right, meaning no
                          security-filtered principal can apply the policy.

    Both result sets are written to a single Excel workbook (two worksheets) when
    the ImportExcel module is available, or to individual CSV files as a fallback.
    A timestamped log file is always written to a 'Log' sub-folder next to the script.

    Output filenames include the domain name and the run date, e.g.:
      2026-06-22_contoso.com_Get-UnusedGPOs.xlsx
.PARAMETER Domain
    FQDN of the Active Directory domain to query.
    Defaults to the current user's logon domain ($env:USERDNSDOMAIN).
.PARAMETER OutputPath
    Directory where the Excel/CSV report is written.
    Defaults to the directory containing this script ($PSScriptRoot).
.OUTPUTS
    <OutputPath>\YYYY-MM-dd_<Domain>_Get-UnusedGPOs.xlsx
      Worksheet 'Unlinked'   — GPOs not linked to any container.
      Worksheet 'NoApplyACE' — GPOs with no Allow 'Apply Group Policy' ACE.
    (or two separate CSV files when ImportExcel is unavailable)
    <PSScriptRoot>\Log\YYYYMMDD_HHmmss_<Domain>_Cleanup-Policies.log
.EXAMPLE
    .\Cleanup-Policies.ps1

    Runs against the current user's domain and writes output to the script directory.
.EXAMPLE
    .\Cleanup-Policies.ps1 -Domain contoso.com -OutputPath C:\Reports

    Runs against contoso.com and writes the report to C:\Reports.
.NOTES
    Author      : M. Stam
    Date        : 2026-06-25
    Version     : 1.1.0

    Requires    : GroupPolicy module (RSAT-GPMC)
                  ActiveDirectory module (RSAT-AD-PowerShell)
                  ImportExcel module (optional; https://github.com/dfinke/ImportExcel)

    Permissions : Domain read access.
                  Group Policy: Read permission on all GPOs.

    Apply ACE detection uses Get-GPPermission -All with the locale-independent
    GPPermissionType enum value 'GpoApply'. The earlier approach of parsing
    Get-GPOReport XML relied on the localised string 'Apply Group Policy'
    (Dutch: 'Groepsbeleid toepassen'), which caused false positives.

    Version history:
      1.1.0  2026-06-25  M. Stam  Fixed Apply ACE detection to use Get-GPPermission
                                   instead of XML report parsing (locale-independent).
                                   Added per-GPO [HasApplyACE]/[NoApplyACE] log entries.
      1.0.0  2026-06-22  M. Stam  Initial release.
#>
#Requires -Modules GroupPolicy, ActiveDirectory

[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $Domain = $env:USERDNSDOMAIN,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputPath = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Logging
$LogDir = Join-Path $PSScriptRoot 'Log'
if (-not (Test-Path $LogDir)) { $null = New-Item -ItemType Directory -Path $LogDir }
$LogFile = Join-Path $LogDir "$(Get-Date -Format 'yyyyMMdd_HHmmss')_${Domain}_Cleanup-Policies.log"

function Write-ScriptLog {
    param ([string] $Message, [string] $Level = 'INFO')
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $entry | Tee-Object -FilePath $LogFile -Append | Write-Verbose
    if ($Level -eq 'ERROR') { Write-Error $Message }
    elseif ($Level -eq 'WARN') { Write-Warning $Message }
    else { Write-Information -MessageData $entry -InformationAction Continue }
}
#endregion

#region Helper – collect all GPO links recursively
function Get-AllGpoLink {
    <#
    .SYNOPSIS Returns a hashtable keyed on GPO ID with all linked container DNs.
    #>
    param ([string] $DomainFqdn)

    $links = @{}  # [Guid] -> [List[string]]

    # Query domain root + all OUs + all sites via Get-ADOrganizationalUnit / Get-ADDomain
    $containers = @()

    # Domain root itself
    $domainDN = (Get-ADDomain -Server $DomainFqdn).DistinguishedName
    $containers += Get-ADObject -Identity $domainDN -Properties 'gpLink' -Server $DomainFqdn

    # All OUs
    $containers += Get-ADOrganizationalUnit -Filter * -Properties 'gpLink' -Server $DomainFqdn

    # Sites (stored in the Configuration partition)
    try {
        $configNC = (Get-ADRootDSE -Server $DomainFqdn).configurationNamingContext
        $containers += Get-ADObject -Filter { objectClass -eq 'site' } `
            -SearchBase $configNC -Properties 'gpLink' -Server $DomainFqdn
    }
    catch {
        Write-ScriptLog "Could not enumerate sites: $_" -Level 'WARN'
    }

    foreach ($container in $containers) {
        if ([string]::IsNullOrWhiteSpace($container.gpLink)) { continue }

        # gpLink format: [LDAP://cn={GUID},cn=Policies,cn=System,DC=...;FLAGS]
        $linkMatches = [regex]::Matches($container.gpLink, '\{(?<guid>[0-9A-Fa-f\-]{36})\}')
        foreach ($m in $linkMatches) {
            $guid = [guid] $m.Groups['guid'].Value
            if (-not $links.ContainsKey($guid)) { $links[$guid] = [System.Collections.Generic.List[string]]::new() }
            $links[$guid].Add($container.DistinguishedName)
        }
    }
    return $links
}
#endregion

#region Main
Write-ScriptLog "Starting GPO inventory for domain: $Domain"

# Collect all GPO links
Write-ScriptLog "Collecting all GPO links..."
$allLinks = Get-AllGpoLink -DomainFqdn $Domain

# Retrieve every GPO
Write-ScriptLog "Retrieving all GPOs..."
$allGpos = Get-GPO -All -Domain $Domain

$unlinked    = [System.Collections.Generic.List[PSCustomObject]]::new()
$noApplyAce  = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-ScriptLog "Analysing $($allGpos.Count) GPOs..."

foreach ($gpo in $allGpos) {
    $id = $gpo.Id   # [Guid]

    #--- Check 1: no links ---
    $linkedTo = if ($allLinks.ContainsKey($id)) { $allLinks[$id] -join '; ' } else { $null }
    $isUnlinked = -not $allLinks.ContainsKey($id)

    #--- Check 2: no 'Apply Group Policy' Allow ACE ---
    # Use Get-GPPermission -All instead of XML report parsing: the GPPermissionType enum
    # value 'GpoApply' is locale-independent, whereas the XML GPOGroupedAccessEnum text
    # is localised (e.g. Dutch: "Groepsbeleid toepassen") and would cause false positives.
    try {
        $perms = Get-GPPermission -Guid $id -All -Domain $Domain -ErrorAction Stop
        $hasApplyAce = [bool]($perms | Where-Object { $_.Permission -eq 'GpoApply' -and -not $_.Denied })
        $aceStatus = if ($hasApplyAce) { 'HasApplyACE' } else { 'NoApplyACE' }
        Write-ScriptLog "  [$aceStatus] $($gpo.DisplayName)"
    }
    catch {
        Write-ScriptLog "Could not read ACL for GPO '$($gpo.DisplayName)': $_" -Level 'WARN'
        $hasApplyAce = $true  # assume OK if we cannot read
    }

    $record = [PSCustomObject]@{
        GPOName       = $gpo.DisplayName
        GPOId         = $id.ToString('B').ToUpper()
        GpoStatus     = $gpo.GpoStatus
        CreationTime  = $gpo.CreationTime
        ModificationTime = $gpo.ModificationTime
        LinkedTo      = $linkedTo
        IsUnlinked    = $isUnlinked
        HasApplyAce   = $hasApplyAce
    }

    if ($isUnlinked)   { $unlinked.Add($record) }
    if (-not $hasApplyAce) { $noApplyAce.Add($record) }
}

Write-ScriptLog "Unlinked GPOs      : $($unlinked.Count)"
Write-ScriptLog "GPOs without Apply ACE: $($noApplyAce.Count)"

#endregion

#region Export
$datePrefix  = Get-Date -Format 'yyyy-MM-dd'
$xlsxPath    = Join-Path $OutputPath "${datePrefix}_${Domain}_Get-UnusedGPOs.xlsx"
$csvPathBase = Join-Path $OutputPath "${datePrefix}_${Domain}_Get-UnusedGPOs"

$hasImportExcel = Get-Module -ListAvailable -Name ImportExcel

if ($hasImportExcel) {
    Write-ScriptLog "Exporting to Excel: $xlsxPath"

    if ($unlinked.Count -gt 0) {
        $unlinked | Export-Excel -Path $xlsxPath -WorksheetName 'Unlinked' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($noApplyAce.Count -gt 0) {
        $noApplyAce | Export-Excel -Path $xlsxPath -WorksheetName 'NoApplyACE' `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }

    if ($unlinked.Count -eq 0 -and $noApplyAce.Count -eq 0) {
        Write-ScriptLog "No unused GPOs found; no output file written." -Level 'WARN'
    }
    else {
        Write-ScriptLog "Report written to: $xlsxPath"
    }
}
else {
    Write-ScriptLog "ImportExcel not available; falling back to CSV." -Level 'WARN'

    if ($unlinked.Count -gt 0) {
        $csvUnlinked = "${csvPathBase}_Unlinked.csv"
        $unlinked | Export-Csv -Path $csvUnlinked -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "Unlinked GPOs CSV: $csvUnlinked"
    }

    if ($noApplyAce.Count -gt 0) {
        $csvNoApply = "${csvPathBase}_NoApplyACE.csv"
        $noApplyAce | Export-Csv -Path $csvNoApply -NoTypeInformation -Encoding UTF8
        Write-ScriptLog "No Apply ACE CSV : $csvNoApply"
    }

    if ($unlinked.Count -eq 0 -and $noApplyAce.Count -eq 0) {
        Write-ScriptLog "No unused GPOs found; no output file written." -Level 'WARN'
    }
}
#endregion
