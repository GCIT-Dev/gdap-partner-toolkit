function Get-MspCustomerConsentState {
    <#
    .SYNOPSIS
        Reads, through Microsoft Graph in the customer tenant, what the partner app actually holds there.
    .DESCRIPTION
        For each desired resource: whether the resource service principal exists in the customer, the
        tenant-wide (AllPrincipals) delegated grant for the partner app, and the missing and extra scopes.
        Also returns the app's grants for resources outside the desired set (OtherGrants, so a consent
        replacement can keep them) and any app role assignments held by the partner app's service principal.
    .PARAMETER TenantId
        Customer tenant ID.
    .PARAMETER AppId
        Partner app ID.
    .PARAMETER Grant
        Desired grants from Get-MspDesiredDelegatedGrant (ResourceAppId, ResourceDisplayName, Scopes).
        Throws when Graph cannot be read, so callers can report Unknown instead of guessing.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Grant
    )
    $appSp = Get-MspServicePrincipalByAppId -TenantId $TenantId -AppId $AppId -Select 'id,appId,displayName,appOwnerOrganizationId'
    $grantsByResource = @{}
    $appRoleAssignments = @()
    if ($appSp) {
        foreach ($existing in @(Invoke-MspGraphCall -TenantId $TenantId -Path ("oauth2PermissionGrants?`$filter=clientId eq '{0}' and consentType eq 'AllPrincipals'" -f $appSp.id))) {
            if ($existing) { $grantsByResource[[string]$existing.resourceId] = $existing }
        }
        $appRoleAssignments = @(Invoke-MspGraphCall -TenantId $TenantId -Path ("servicePrincipals/{0}/appRoleAssignments" -f $appSp.id))
    }

    $desiredSpIds = New-Object System.Collections.Generic.List[string]
    $resources = foreach ($desired in $Grant) {
        $resourceSp = Get-MspServicePrincipalByAppId -TenantId $TenantId -AppId $desired.ResourceAppId
        if ($resourceSp) { $desiredSpIds.Add([string]$resourceSp.id) }
        $existingGrant = if ($resourceSp -and $grantsByResource.ContainsKey([string]$resourceSp.id)) { $grantsByResource[[string]$resourceSp.id] } else { $null }
        $current = if ($existingGrant) { @(([string]$existingGrant.scope).Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) } else { @() }
        [pscustomobject]@{
            ResourceAppId       = $desired.ResourceAppId
            ResourceDisplayName = $desired.ResourceDisplayName
            ResourcePresent     = [bool]$resourceSp
            ResourceSpId        = if ($resourceSp) { [string]$resourceSp.id } else { $null }
            GrantId             = if ($existingGrant) { [string]$existingGrant.id } else { $null }
            DesiredScopes       = @($desired.Scopes)
            CurrentScopes       = $current
            MissingScopes       = @($desired.Scopes | Where-Object { $current -notcontains $_ })
            ExtraScopes         = @($current | Where-Object { @($desired.Scopes) -notcontains $_ })
        }
    }

    # Grants the app holds for resources outside the desired set (excluded or not in the manifest).
    $otherGrants = foreach ($key in @($grantsByResource.Keys)) {
        if ($desiredSpIds -contains $key) { continue }
        [pscustomobject]@{
            ResourceSpId = $key
            GrantId      = [string]$grantsByResource[$key].id
            Scopes       = @(([string]$grantsByResource[$key].scope).Split(' ', [StringSplitOptions]::RemoveEmptyEntries))
        }
    }

    [pscustomobject]@{
        TenantId            = $TenantId
        AppServicePrincipal = $appSp
        Resources           = @($resources)
        OtherGrants         = @($otherGrants)
        AppRoleAssignments  = $appRoleAssignments
    }
}
