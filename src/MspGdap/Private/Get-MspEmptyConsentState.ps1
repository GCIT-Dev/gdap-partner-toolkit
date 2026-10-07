function Get-MspEmptyConsentState {
    <#
    .SYNOPSIS
        A consent state for a customer where the partner app has never been consented.
    .DESCRIPTION
        Same shape as Get-MspCustomerConsentState. Every desired resource is assumed present (Partner Center
        reports a resource that is missing in the customer when it is posted) and every desired scope is missing.
    .PARAMETER TenantId
        Customer tenant ID.
    .PARAMETER Grant
        Desired grants from Get-MspDesiredDelegatedGrant (ResourceAppId, ResourceDisplayName, Scopes).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Grant
    )
    $resources = foreach ($desired in $Grant) {
        [pscustomobject]@{
            ResourceAppId       = $desired.ResourceAppId
            ResourceDisplayName = $desired.ResourceDisplayName
            ResourcePresent     = $true
            ResourceSpId        = $null
            GrantId             = $null
            DesiredScopes       = @($desired.Scopes)
            CurrentScopes       = @()
            MissingScopes       = @($desired.Scopes)
            ExtraScopes         = @()
        }
    }
    [pscustomobject]@{
        TenantId            = $TenantId
        AppServicePrincipal = $null
        Resources           = @($resources)
        OtherGrants         = @()
        AppRoleAssignments  = @()
        NotConsented        = $true
    }
}
