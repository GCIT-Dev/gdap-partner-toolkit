function Get-MspTokenCacheKey {
    <#
    .SYNOPSIS
        Returns the token cache key "<tenantId>|<resource>|<appId>" in lower case.
    .DESCRIPTION
        Access tokens are single-audience and tenant specific, and refresh
        tokens are bound to the client, so the key carries all three. The
        resource is normalised (alias resolved, trailing slash and /.default
        removed) so that equivalent requests share a cache entry.
    .PARAMETER TenantId
        Tenant GUID the token was issued for.
    .PARAMETER Resource
        Resource alias, URI or application ID.
    .PARAMETER AppId
        Partner app (client) ID.
    .EXAMPLE
        Get-MspTokenCacheKey -TenantId $tid -Resource Graph -AppId $appId
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$Resource,

        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$AppId
    )
    $resourceUri = (Resolve-MspResource -Resource $Resource).Uri
    ('{0}|{1}|{2}' -f $TenantId, $resourceUri, $AppId).ToLowerInvariant()
}
