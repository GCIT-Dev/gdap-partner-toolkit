function Invoke-MspRefreshTokenRedemption {
    <#
    .SYNOPSIS
        Redeems the technician refresh token at a tenant's token endpoint and
        writes the rotated refresh token back to the vault.
    .DESCRIPTION
        Steps:
        1. Read the refresh token from SecretManagement.
        2. POST grant_type=refresh_token to /<TenantId>/oauth2/v2.0/token (the
           CUSTOMER tenant for customer work, so GDAP decides the access).
        3. If the token can be decoded and its tid is not the requested
           tenant, refuse it (wrong-tenant protection).
        4. Write the new refresh token back to the vault. A failed write-back
           is a warning, not a failure, because the previous token stays valid
           until it expires.
        5. Return a cache entry with the access token as a SecureString and the
           expiry taken from expires_in in the token response.
    .PARAMETER Configuration
        The active configuration object.
    .PARAMETER TenantId
        Tenant GUID to redeem against.
    .PARAMETER Request
        Output of Resolve-MspTokenRequestScope.
    .PARAMETER UserPrincipalName
        Technician UPN.
    .EXAMPLE
        Invoke-MspRefreshTokenRedemption -Configuration $config -TenantId $tid -Request $request -UserPrincipalName $upn
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Configuration,

        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [pscustomobject]$Request,

        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $refreshToken = Get-MspRefreshToken -UserPrincipalName $UserPrincipalName -AppId $Configuration.AppId -VaultName $Configuration.VaultName

    $body = [ordered]@{
        client_id     = $Configuration.AppId
        grant_type    = 'refresh_token'
        scope         = $Request.ScopeString
        refresh_token = ConvertFrom-MspSecureString -SecureString $refreshToken
    }
    $credential = Get-MspClientCredentialBody -Configuration $Configuration -TenantId $TenantId
    foreach ($key in $credential.Keys) { $body[$key] = $credential[$key] }

    $requestedAt = [DateTimeOffset]::UtcNow
    try {
        $response = Invoke-MspTokenRequest -TenantId $TenantId -Body $body
    }
    finally {
        $body = $null
        $refreshToken = $null
    }

    if (-not $response -or -not $response.PSObject.Properties['access_token'] -or -not $response.access_token) {
        throw (New-MspErrorRecord -Message "The token endpoint for tenant $TenantId returned no access token." -ErrorId 'MspGdap.TokenRequest.NoAccessToken' -Category InvalidResult -TargetObject $TenantId)
    }

    $decoded = ConvertFrom-MspJwt -Token $response.access_token
    $claims = if ($decoded) { $decoded.Claims } else { $null }
    if ($claims -and $claims.PSObject.Properties['tid'] -and $claims.tid -and ([string]$claims.tid).ToLowerInvariant() -ne $TenantId.ToLowerInvariant()) {
        throw (New-MspErrorRecord -Message "The token issued for tenant $TenantId carries tenant $($claims.tid). It was discarded." -ErrorId 'MspGdap.Token.TenantMismatch' -Category SecurityError -TargetObject $TenantId)
    }

    if ($response.PSObject.Properties['refresh_token'] -and $response.refresh_token) {
        try {
            $null = Save-MspRefreshToken -UserPrincipalName $UserPrincipalName -AppId $Configuration.AppId -VaultName $Configuration.VaultName -RefreshToken (ConvertTo-MspSecureString -Value $response.refresh_token)
            Write-Verbose "Rotated refresh token for $UserPrincipalName written back to vault $($Configuration.VaultName)."
        }
        catch {
            Write-Warning "The rotated refresh token for $UserPrincipalName could not be written back to vault '$($Configuration.VaultName)'. The previous token remains valid until it expires, but fix vault access soon. $($_.Exception.Message)"
        }
    }
    else {
        Write-Verbose 'The token response did not include a rotated refresh token.'
    }

    $expiresIn = 0
    if ($response.PSObject.Properties['expires_in'] -and $response.expires_in) {
        $expiresIn = [int]$response.expires_in
    }
    elseif ($claims -and $claims.PSObject.Properties['exp'] -and $claims.exp) {
        $expiresIn = [int]([DateTimeOffset]::FromUnixTimeSeconds([long]$claims.exp) - $requestedAt).TotalSeconds
    }

    $scopes = @()
    if ($response.PSObject.Properties['scope'] -and $response.scope) {
        $scopes = @(([string]$response.scope).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { Get-MspScopeShortName -Scope $_ })
    }
    $audience = if ($claims -and $claims.PSObject.Properties['aud']) { [string]$claims.aud } else { $null }

    $entry = [pscustomobject]@{
        PSTypeName        = 'MspGdap.TokenCacheEntry'
        TenantId          = $TenantId.ToLowerInvariant()
        Resource          = $Request.Resource
        AppId             = ([string]$Configuration.AppId).ToLowerInvariant()
        UserPrincipalName = $UserPrincipalName
        AccessToken       = ConvertTo-MspSecureString -Value $response.access_token
        ExpiresOn         = $requestedAt.AddSeconds($expiresIn)
        AcquiredOn        = $requestedAt
        Scopes            = $scopes
        Audience          = $audience
    }
    $response = $null
    $entry
}
