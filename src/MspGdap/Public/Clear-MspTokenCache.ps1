function Clear-MspTokenCache {
    <#
    .SYNOPSIS
        Removes access tokens from the in-memory cache.
    .DESCRIPTION
        With no parameters every cached access token is removed. -TenantId and
        -Resource narrow the removal. Refresh tokens in the vault are not
        touched.
    .PARAMETER TenantId
        Only remove tokens for this tenant (GUID or verified domain).
    .PARAMETER Resource
        Only remove tokens for this resource alias, URI or application ID.
    .EXAMPLE
        Clear-MspTokenCache
    .EXAMPLE
        Clear-MspTokenCache -TenantId 'fabrikam.onmicrosoft.com' -Resource Exchange
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [string]$Resource
    )
    $tenantFilter = if ($TenantId) { Resolve-MspTenantId -Tenant $TenantId } else { $null }
    $resourceUri = if ($Resource) { (Resolve-MspResource -Resource $Resource).Uri } else { $null }
    $removed = 0
    foreach ($key in @($script:MspTokenCache.Keys)) {
        $parts = $key.Split('|')
        if ($tenantFilter -and $parts[0] -ne $tenantFilter) { continue }
        if ($resourceUri -and $parts[1] -ne $resourceUri) { continue }
        $script:MspTokenCache.Remove($key)
        $removed++
    }
    Write-Verbose "Removed $removed cached access token(s)."
}
