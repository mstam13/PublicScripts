@{
    ModuleVersion     = '1.0.0'
    GUID              = '03d050ae-a770-4d27-95a0-dbbb27602759'
    Author            = 'M. Stam'
    Description       = 'Shared logging utilities (Initialize-ScriptLog, Write-ScriptLog, Remove-OldLogs) for PublicScripts repository scripts.'
    PowerShellVersion = '5.1'
    RootModule        = 'PublicScripts.psm1'
    FunctionsToExport = @('Initialize-ScriptLog', 'Write-ScriptLog', 'Remove-OldLog')
    CmdletsToExport   = @()
    AliasesToExport   = @()
    VariablesToExport = @()
}
