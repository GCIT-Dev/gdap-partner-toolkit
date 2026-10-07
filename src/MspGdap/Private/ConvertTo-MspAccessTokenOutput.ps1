function ConvertTo-MspAccessTokenOutput {
    <#
    .SYNOPSIS
        Shapes a cache entry into the object returned by Get-MspAccessToken.
    .PARAMETER Entry
        A cache entry.
    .PARAMETER FromCache
        Whether the entry came from the cache.
    .EXAMPLE
        ConvertTo-MspAccessTokenOutput -Entry $entry -FromCache
    #>
    [CmdletBinding()]
    [OutputType('MspGdap.AccessToken')]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Entry,

        [switch]$FromCache
    )
    [pscustomobject]@{
        PSTypeName        = 'MspGdap.AccessToken'
        TenantId          = $Entry.TenantId
        Resource          = $Entry.Resource
        AppId             = $Entry.AppId
        UserPrincipalName = $Entry.UserPrincipalName
        ExpiresOn         = $Entry.ExpiresOn
        Scopes            = $Entry.Scopes
        FromCache         = [bool]$FromCache
        AccessToken       = $Entry.AccessToken
    }
}
