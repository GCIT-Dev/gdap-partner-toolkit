function Test-MspExoCompatibility {
    <#
    .SYNOPSIS
        Checks an ExchangeOnlineManagement version against the running PowerShell.
    .DESCRIPTION
        Microsoft Learn: module 3.10.0 and later require PowerShell 7.6.0 or later (Windows PowerShell 5.1 is not
        affected). Modules 3.5.0 to 3.9.2 require PowerShell 7.4.0 or later on PowerShell 7.
        -AccessToken on Connect-ExchangeOnline needs 3.1.0 or later. -AccessToken on Connect-IPPSSession needs 3.8.0 or later.
        Returns $null when compatible, otherwise a message explaining what to change.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][version]$ModuleVersion,
        [version]$PSVersion = $PSVersionTable.PSVersion,
        [string]$Edition = $PSVersionTable.PSEdition,
        [version]$MinimumModuleVersion = '3.1.0'
    )
    if ($ModuleVersion -lt $MinimumModuleVersion) {
        return "ExchangeOnlineManagement $ModuleVersion is too old. $MinimumModuleVersion or later is required for token-based connections."
    }
    if ($Edition -eq 'Core') {
        if ($ModuleVersion -ge [version]'3.10.0' -and $PSVersion -lt [version]'7.6.0') {
            return "ExchangeOnlineManagement $ModuleVersion needs PowerShell 7.6 or later (you are on $PSVersion). Upgrade PowerShell, or install ExchangeOnlineManagement 3.9.2 and import it with -RequiredVersion 3.9.2."
        }
        if ($ModuleVersion -ge [version]'3.5.0' -and $PSVersion -lt [version]'7.4.0') {
            return "ExchangeOnlineManagement $ModuleVersion needs PowerShell 7.4 or later (you are on $PSVersion)."
        }
    }
    return $null
}
