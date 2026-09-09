@{
    ModuleVersion     = '1.1.0'
    GUID              = '03d050ae-a770-4d27-95a0-dbbb27602759'
    Author            = 'M. Stam'
    Description       = 'Shared logging (Initialize-ScriptLog, Write-ScriptLog, Remove-OldLog), module-import (Import-RequiredModule), and GPO XML report parsing (Get-XmlInnerText, Get-GpoSetting) utilities for PublicScripts repository scripts.'
    PowerShellVersion = '5.1'
    RootModule        = 'PublicScripts.psm1'
    FunctionsToExport = @('Initialize-ScriptLog', 'Write-ScriptLog', 'Remove-OldLog', 'Import-RequiredModule', 'Get-XmlInnerText', 'Get-GpoSetting')
    CmdletsToExport   = @()
    AliasesToExport   = @()
    VariablesToExport = @()
}
