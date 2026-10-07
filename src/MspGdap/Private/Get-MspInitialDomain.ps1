function Get-MspInitialDomain {
    <#
    .SYNOPSIS
        Returns the initial .onmicrosoft.com domain of a customer tenant, and checks the token really is for that tenant.
    .DESCRIPTION
        Reads GET /organization?$select=id,displayName,verifiedDomains with the customer token and picks the
        verified domain where isInitial is true. Throws if the organisation ID differs from -TenantId, which
        would mean the call reached the wrong tenant.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$TenantId)

    $organization = @(Invoke-MspGraphCall -TenantId $TenantId -Path 'organization?$select=id,displayName,verifiedDomains') | Select-Object -First 1
    if (-not $organization) { throw "Could not read the organisation of tenant $TenantId." }
    if ($organization.id -and ($organization.id -ne $TenantId)) {
        throw ("Wrong tenant: asked for {0} but the token returned organisation {1}. Nothing was changed." -f $TenantId, $organization.id)
    }
    $initial = @($organization.verifiedDomains | Where-Object { $_.isInitial }) | Select-Object -First 1
    if (-not $initial) { throw "Tenant $TenantId has no initial domain in verifiedDomains." }
    [pscustomobject]@{
        TenantId      = $TenantId
        DisplayName   = $organization.displayName
        InitialDomain = [string]$initial.name
    }
}
