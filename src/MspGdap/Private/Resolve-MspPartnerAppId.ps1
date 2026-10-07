function Resolve-MspPartnerAppId {
    <#
    .SYNOPSIS
        Returns -AppId when given, otherwise the partner app ID from Get-MspConfiguration. Never guesses.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$AppId)

    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    if ($AppId) {
        if ($AppId -notmatch $guid) { throw "AppId '$AppId' is not a GUID." }
        return $AppId.ToLowerInvariant()
    }
    $configured = $null
    if (Get-Command -Name 'Get-MspConfiguration' -ErrorAction SilentlyContinue) {
        try {
            $config = Get-MspConfiguration -ErrorAction Stop
            if ($config -and $config.PSObject.Properties['AppId']) { $configured = [string]$config.AppId }
        }
        catch { $configured = $null }
    }
    if ($configured -and $configured -match $guid) { return $configured.ToLowerInvariant() }
    throw 'No partner app ID. Pass -AppId or run Set-MspConfiguration -AppId first.'
}
