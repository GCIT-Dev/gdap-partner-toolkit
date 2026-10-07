function Get-MspServicePrincipalByAppId {
    <#
    .SYNOPSIS
        Finds one service principal by appId in a customer tenant or the partner tenant. Returns $null when absent.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant')][string]$TenantId,
        [Parameter(Mandatory, ParameterSetName = 'Partner')][switch]$PartnerTenant,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [string]$Select = 'id,appId,displayName'
    )
    $target = @{}
    if ($PartnerTenant) { $target['PartnerTenant'] = $true } else { $target['TenantId'] = $TenantId }
    $path = "servicePrincipals?`$filter=appId eq '{0}'&`$select={1}" -f $AppId, $Select
    @(Invoke-MspGraphCall @target -Path $path) | Select-Object -First 1
}
