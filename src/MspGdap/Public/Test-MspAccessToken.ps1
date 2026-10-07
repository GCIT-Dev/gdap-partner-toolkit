function Test-MspAccessToken {
    <#
    .SYNOPSIS
        Checks that an access token is still usable for a tenant and resource.
    .DESCRIPTION
        Microsoft treats access tokens as opaque to clients, and some are not
        decodable. The primary checks therefore use the token response
        metadata that MspGdap records when the token is issued:

        - ExpiresOn (from expires_in) must be more than SkewMinutes away.
        - The tenant the token was requested for must equal -TenantId.
        - The resource it was requested for must equal -Resource.
        - Required scopes must be among the scopes the response reported.

        When the token is a decodable JWT the same checks are repeated on the
        claims as a best-effort extra: exp (with skew), tid, aud (the URI or
        application ID form of the resource are both accepted), scp and roles.
        A token that cannot be decoded never fails because of that. A token
        with no known expiry at all is treated as not usable.

        No signature validation is done. This function decides whether to reuse
        a token MspGdap obtained itself. It is not a way to trust a token from
        somewhere else.
    .PARAMETER InputObject
        A cache entry or the output of Get-MspAccessToken.
    .PARAMETER AccessToken
        A raw access token (string or SecureString). Use with -ExpiresOn when
        the token may be opaque.
    .PARAMETER TenantId
        The tenant the token must belong to (GUID or verified domain). Mandatory.
    .PARAMETER Resource
        Resource alias, URI or application ID the token must be for.
    .PARAMETER ExpiresOn
        Expiry from the token response, for raw tokens.
    .PARAMETER RequiredScope
        Delegated scopes that must be present.
    .PARAMETER RequiredRole
        Application roles that must be present (decodable app-only tokens).
    .PARAMETER SkewMinutes
        Minutes before expiry at which a token counts as expired. Default 5.
    .PARAMETER Detailed
        Return an object with IsValid and Reasons instead of a Boolean.
    .EXAMPLE
        Get-MspAccessToken -TenantId $tid | Test-MspAccessToken -TenantId $tid -Resource Graph
    .EXAMPLE
        Test-MspAccessToken -AccessToken $raw -TenantId $tid -Resource 'https://outlook.office365.com' -ExpiresOn $expiry -Detailed
    #>
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    [OutputType([bool])]
    [OutputType('MspGdap.TokenTestResult')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'InputObject', ValueFromPipeline)]
        [pscustomobject]$InputObject,

        [Parameter(Mandatory, ParameterSetName = 'Token', Position = 0)]
        [object]$AccessToken,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [string]$Resource,

        [Parameter(ParameterSetName = 'Token')]
        [Nullable[DateTimeOffset]]$ExpiresOn,

        [string[]]$RequiredScope,

        [string[]]$RequiredRole,

        [ValidateRange(0, 60)]
        [int]$SkewMinutes = 5,

        [switch]$Detailed
    )

    process {
        $reasons = [System.Collections.Generic.List[string]]::new()
        $expectedTenant = Resolve-MspTenantId -Tenant $TenantId
        $resourceInfo = if ($Resource) { Resolve-MspResource -Resource $Resource } else { $null }

        $token = $null
        $metaExpiry = $null
        $metaTenant = $null
        $metaResource = $null
        $metaScopes = @()
        $metaAudience = $null

        if ($PSCmdlet.ParameterSetName -eq 'InputObject') {
            $properties = $InputObject.PSObject.Properties
            if ($properties['AccessToken']) { $token = $InputObject.AccessToken }
            if ($properties['ExpiresOn'] -and $InputObject.ExpiresOn) { $metaExpiry = [DateTimeOffset]$InputObject.ExpiresOn }
            if ($properties['TenantId'] -and $InputObject.TenantId) { $metaTenant = ([string]$InputObject.TenantId).ToLowerInvariant() }
            if ($properties['Resource'] -and $InputObject.Resource) { $metaResource = (Resolve-MspResource -Resource $InputObject.Resource).Uri }
            if ($properties['Scopes'] -and $InputObject.Scopes) { $metaScopes = @($InputObject.Scopes) }
            if ($properties['Audience'] -and $InputObject.Audience) { $metaAudience = [string]$InputObject.Audience }
        }
        else {
            $token = $AccessToken
            if ($ExpiresOn) { $metaExpiry = [DateTimeOffset]$ExpiresOn }
        }

        if ($null -eq $token -or ($token -is [string] -and [string]::IsNullOrWhiteSpace($token))) {
            $reasons.Add('NoToken')
        }

        $decoded = if ($null -ne $token) { ConvertFrom-MspJwt -Token $token } else { $null }
        $claims = if ($decoded) { $decoded.Claims } else { $null }
        $threshold = [DateTimeOffset]::UtcNow.AddMinutes($SkewMinutes)

        # Expiry: metadata first, claims as an extra.
        $effectiveExpiry = $metaExpiry
        if ($claims -and $claims.PSObject.Properties['exp'] -and $claims.exp) {
            $claimExpiry = [DateTimeOffset]::FromUnixTimeSeconds([long]$claims.exp)
            if ($null -eq $effectiveExpiry -or $claimExpiry -lt $effectiveExpiry) { $effectiveExpiry = $claimExpiry }
        }
        if ($null -eq $effectiveExpiry) {
            $reasons.Add('ExpiryUnknown')
        }
        elseif ($effectiveExpiry -le $threshold) {
            $reasons.Add('Expired')
        }

        # Tenant.
        if ($metaTenant -and $metaTenant -ne $expectedTenant) { $reasons.Add('TenantMismatch') }
        if ($claims -and $claims.PSObject.Properties['tid'] -and $claims.tid -and ([string]$claims.tid).ToLowerInvariant() -ne $expectedTenant) {
            if (-not $reasons.Contains('TenantMismatch')) { $reasons.Add('TenantMismatch') }
        }

        # Resource and audience.
        if ($resourceInfo) {
            if ($metaResource -and $metaResource -ne $resourceInfo.Uri) { $reasons.Add('ResourceMismatch') }
            if ($claims -and $claims.PSObject.Properties['aud'] -and $claims.aud) {
                $aud = ([string]$claims.aud).ToLowerInvariant()
                $accepted = @($resourceInfo.Audiences | ForEach-Object { $_.ToLowerInvariant() })
                if ($metaAudience -and $metaResource -eq $resourceInfo.Uri) { $accepted += $metaAudience.ToLowerInvariant() }
                if ($accepted -notcontains $aud -and $accepted -notcontains $aud.TrimEnd('/')) {
                    $reasons.Add('AudienceMismatch')
                }
            }
        }

        # Scopes and roles.
        if ($RequiredScope) {
            $available = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($s in $metaScopes) { $null = $available.Add((Get-MspScopeShortName -Scope $s)) }
            if ($claims -and $claims.PSObject.Properties['scp'] -and $claims.scp) {
                foreach ($s in ([string]$claims.scp).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)) { $null = $available.Add($s) }
            }
            if ($available.Count -eq 0) {
                $reasons.Add('ScopesUnverifiable')
            }
            else {
                foreach ($s in $RequiredScope) {
                    if (-not $available.Contains((Get-MspScopeShortName -Scope $s))) { $reasons.Add("MissingScope:$s") }
                }
            }
        }
        if ($RequiredRole) {
            $roles = if ($claims -and $claims.PSObject.Properties['roles']) { @($claims.roles) } else { @() }
            if ($roles.Count -eq 0) {
                $reasons.Add('RolesUnverifiable')
            }
            else {
                foreach ($r in $RequiredRole) {
                    if ($roles -notcontains $r) { $reasons.Add("MissingRole:$r") }
                }
            }
        }

        $isValid = $reasons.Count -eq 0
        Write-Verbose ("Token check for tenant {0}: {1}{2}" -f $expectedTenant, $(if ($isValid) { 'valid' } else { 'not valid' }), $(if ($isValid) { '' } else { " ($($reasons -join ', '))" }))

        if ($Detailed) {
            [pscustomobject]@{
                PSTypeName    = 'MspGdap.TokenTestResult'
                IsValid       = $isValid
                Reasons       = $reasons.ToArray()
                ExpiresOn     = $effectiveExpiry
                TenantId      = $expectedTenant
                Resource      = if ($resourceInfo) { $resourceInfo.Uri } else { $metaResource }
                ClaimsChecked = [bool]$claims
            }
        }
        else {
            $isValid
        }
    }
}
