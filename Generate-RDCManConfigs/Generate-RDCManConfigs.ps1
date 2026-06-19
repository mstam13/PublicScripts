#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Generates RDCMan (.rdg) configuration files for each Active Directory domain.

.DESCRIPTION
    Automates the creation of Remote Desktop Connection Manager configuration files for
    one or more AD domains. For each domain, the script queries all enabled Windows
    Server computer objects and organises them into nested groups that mirror their
    OU/canonical-path structure. The result is saved as a .rdg file readable by
    RDCMan 2.93+.

    All actions are logged to a timestamped log file under .\Log\.

.PARAMETER RDCManPath
    Directory path where the generated RDCMan configuration (.rdg) files will be saved.
    The directory must already exist.

.PARAMETER DomainNames
    One or more fully qualified domain names (FQDNs) to query. Only one domain
    controller FQDN per domain is required.

.PARAMETER UseCurrentCredentials
    When specified, uses the credentials of the running account instead of prompting
    for credentials per domain. Useful for unattended or scheduled execution.

.PARAMETER Force
    Overwrites existing .rdg files without warning. Without this switch the script
    skips any domain whose output file already exists.

.OUTPUTS
    One .rdg file per domain saved to RDCManPath.
    Log file written to .\Log\<timestamp>.log.

.EXAMPLE
    .\Generate-RDCManConfigs.ps1 -RDCManPath C:\RDCConfigs -DomainNames contoso.com, fabrikam.com

    Prompts for credentials per domain and generates RDCMan configuration files.

.EXAMPLE
    .\Generate-RDCManConfigs.ps1 -RDCManPath C:\RDCConfigs -DomainNames contoso.com -UseCurrentCredentials -Force

    Runs unattended using the current account's credentials, overwriting any existing output.

.EXAMPLE
    .\Generate-RDCManConfigs.ps1 -RDCManPath C:\RDCConfigs -DomainNames contoso.com -WhatIf

    Shows what would be saved without writing any files.

.NOTES
    Author:   Joey Eckelbarger
    Editor:   Marcel Stam
    Version:  2.0.0

    Version History:
    1.0.0 - Joey Eckelbarger  - Initial release
    2.0.0 - Marcel Stam       - 2026-06-17
                                  Added param block with ValidateScript
                                  Fixed: -Credential now passed to Get-ADComputer
                                  Added -UseCurrentCredentials switch for unattended use
                                  Added -Force switch to control overwrite behaviour
                                  Added SupportsShouldProcess / -WhatIf support
                                  Wrapped AD queries and file saves in try/catch
                                  Added Write-ScriptLog to .\.Log\ with timestamp
                                  XPath injection protection via ConvertTo-XPathLiteral
                                  Replaced shared XML templates with ConvertTo-RDCGroupNode
                                  and ConvertTo-RDCServerNode helpers (no template mutation)
                                  Merged OU-tagging loop into XML-building loop
                                  Replaced string concatenation with Join-Path
                                  Removed Select alias; removed dead variables
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$RDCManPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$DomainNames,

    [Parameter()]
    [switch]$UseCurrentCredentials,

    [Parameter()]
    [switch]$Force
)

#region Logging setup
$script:LogFile = Join-Path $PSScriptRoot "Log\$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$null = New-Item -ItemType Directory -Path (Split-Path $script:LogFile) -Force

function Write-ScriptLog {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Write-Output $entry
    Add-Content -Path $script:LogFile -Value $entry
}
#endregion

#region Helper functions

# Converts a string to a safe XPath 1.0 literal, handling embedded quote characters.
function ConvertTo-XPathLiteral {
    param([string]$Value)
    if ($Value -notmatch "'") { return "'$Value'" }
    if ($Value -notmatch '"') { return '"' + $Value + '"' }
    # Value contains both quote types — use XPath concat() to safely compose the literal.
    $sq    = "'"  # single-quote character
    $glue  = ',"' + $sq + '",'
    $parts = ($Value -split $sq) | ForEach-Object { $sq + $_ + $sq }
    return 'concat(' + ($parts -join $glue) + ')'
}

# Creates a fresh <group> node imported into the target XmlDocument.
# Using a helper avoids mutating a shared template across iterations.
function ConvertTo-RDCGroupNode {
    param(
        [System.Xml.XmlDocument]$Document,
        [string]$Name
    )
    [xml]$template = '<group><properties><expanded>False</expanded><name>X</name></properties></group>'
    $node = $Document.ImportNode($template.DocumentElement, $true)
    $node.properties.name = $Name
    return $node
}

# Creates a fresh <server> node imported into the target XmlDocument.
function ConvertTo-RDCServerNode {
    param(
        [System.Xml.XmlDocument]$Document,
        [string]$DisplayName,
        [string]$ServerFQDN
    )
    [xml]$template = '<server><properties><displayName>X</displayName><name>X</name></properties></server>'
    $node = $Document.ImportNode($template.DocumentElement, $true)
    $node.properties.displayName = $DisplayName
    $node.properties.name        = $ServerFQDN
    return $node
}
#endregion

#region Process each domain
foreach ($domain in $DomainNames) {

    $domain = $domain.Trim()
    Write-ScriptLog "Processing domain: $domain"

    #region Credentials
    if ($UseCurrentCredentials) {
        $adParams = @{ Server = $domain }
        Write-ScriptLog "Using current credentials for '$domain'."
    } else {
        $Credential = Get-Credential -Message "Enter credentials for $domain.`n`nFormat: DOMAIN\USERNAME or USERNAME@DOMAIN.TLD"
        $adParams = @{ Server = $domain; Credential = $Credential }
    }
    #endregion

    #region Build per-domain RDCMan XML
    [xml]$RDCManConfigurationXML = @'
<?xml version="1.0" encoding="utf-8"?>
<RDCMan programVersion="2.93" schemaVersion="3">
    <file>
        <credentialsProfiles />
        <properties>
            <expanded>True</expanded>
            <name>CONFIGNAME</name>
        </properties>
    </file>
    <connected />
    <favorites />
    <recentlyUsed />
</RDCMan>
'@
    #endregion

    #region Query Active Directory
    try {
        $servers = Get-ADComputer `
            -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' `
            -Properties Name, CanonicalName, OperatingSystem `
            @adParams |
            Sort-Object CanonicalName
    }
    catch {
        Write-ScriptLog "ERROR querying AD for domain '$domain': $_"
        continue
    }

    if (-not $servers -or $servers.Count -eq 0) {
        Write-ScriptLog "No servers found in '$domain'. Skipping."
        continue
    }
    #endregion

    $domainName = $servers[0].CanonicalName.Split('/')[0]
    Write-ScriptLog "Found $($servers.Count) servers in $domainName"

    $RDCManConfigurationXML.RDCMan.file.properties.name = $domainName

    #region Build XML tree
    foreach ($server in $servers) {
        $foundNode                 = @()
        $remainingSubfoldersNeeded = $null

        $name        = $server.Name
        [array]$path = $server.CanonicalName.Replace("$domainName/", '').Replace("/$name", '').Split('/')

        # Build an XPath expression to locate existing group nodes for this OU path.
        # ConvertTo-XPathLiteral ensures OU names containing quotes do not break the expression.
        $nodeXpathFilterArray = @()
        foreach ($folder in $path) {
            $literal = ConvertTo-XPathLiteral $folder
            $nodeXpathFilterArray += if ($folder -eq $path[-1]) {
                "//properties[name=$literal]"
            } else {
                "//group[properties[name=$literal]]"
            }

            $nodeXpathFilter = $nodeXpathFilterArray -join ''
            $node = $RDCManConfigurationXML.SelectSingleNode($nodeXpathFilter)

            if ($node) {
                $foundNode += $node
            } elseif ($foundNode) {
                $start = $path.IndexOf($folder)
                $end   = $path.Count - 1
                $remainingSubfoldersNeeded = $path[$start..$end]
                break
            }
        }

        if ($remainingSubfoldersNeeded -or $foundNode) {
            $path       = $remainingSubfoldersNeeded
            $parentNode = $foundNode[-1]

            # Determine anchor node within the found parent
            if ($parentNode.server) {
                $anchorNode = $parentNode.server | Select-Object -Last 1
            } elseif ($parentNode.group) {
                $anchorNode = $parentNode.group | Select-Object -Last 1
            } else {
                $anchorNode = $parentNode.ParentNode.properties
                $parentNode = $parentNode.ParentNode
            }
        } else {
            # No existing node found — insert under the file root
            $parentNode = $RDCManConfigurationXML.RDCMan.file
            if ($RDCManConfigurationXML.RDCMan.file.group) {
                $anchorNode = $parentNode.group | Select-Object -Last 1
            } else {
                $anchorNode = $parentNode.properties
            }
        }

        # Create any missing intermediate group nodes using a fresh node per iteration
        if ($remainingSubfoldersNeeded -or -not $foundNode) {
            foreach ($subfolder in $path) {
                $importedNode = ConvertTo-RDCGroupNode -Document $RDCManConfigurationXML -Name $subfolder
                # Nest subfolders by updating parent/anchor on each iteration
                $parentNode = $parentNode.InsertAfter($importedNode, $anchorNode)
                $anchorNode = $parentNode.properties
            }
        }

        # Insert the server node using a fresh node (avoids mutating a shared template)
        $importedNode = ConvertTo-RDCServerNode -Document $RDCManConfigurationXML `
            -DisplayName $name -ServerFQDN "$name.$domainName"
        [void]$parentNode.InsertAfter($importedNode, $anchorNode)
    }
    #endregion

    #region Save output file
    $outputFile = Join-Path $RDCManPath "$domainName.rdg"

    if ((Test-Path $outputFile) -and -not $Force) {
        Write-ScriptLog "WARNING: '$outputFile' already exists. Use -Force to overwrite. Skipping."
        continue
    }

    Write-ScriptLog "Saving $outputFile ..."
    if ($PSCmdlet.ShouldProcess($outputFile, 'Save RDCMan configuration')) {
        try {
            $RDCManConfigurationXML.Save($outputFile)
            Write-ScriptLog "Saved successfully."
        }
        catch {
            Write-ScriptLog "ERROR saving '$outputFile': $_"
        }
    }
    #endregion
}
#endregion
