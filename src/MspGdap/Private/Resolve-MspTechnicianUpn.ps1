function Resolve-MspTechnicianUpn {
    <#
    .SYNOPSIS
        Picks the technician whose refresh token is used.
    .DESCRIPTION
        Order: explicit -UserPrincipalName, then the technician registered in
        this session, then TechnicianUpn from the configuration file.
    .PARAMETER UserPrincipalName
        Optional explicit UPN.
    .PARAMETER Configuration
        The active configuration object.
    .EXAMPLE
        Resolve-MspTechnicianUpn -Configuration $config
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [pscustomobject]$Configuration
    )
    $upn = if ($UserPrincipalName) {
        $UserPrincipalName
    }
    elseif ($script:MspState.CurrentUpn) {
        $script:MspState.CurrentUpn
    }
    else {
        $Configuration.TechnicianUpn
    }
    if ([string]::IsNullOrWhiteSpace($upn)) {
        throw (New-MspErrorRecord -Message 'No technician is selected. Run Register-MspPartnerToken, pass -UserPrincipalName, or set Set-MspConfiguration -TechnicianUpn.' -ErrorId 'MspGdap.Technician.Missing' -Category InvalidArgument)
    }
    $upn.Trim().ToLowerInvariant()
}
