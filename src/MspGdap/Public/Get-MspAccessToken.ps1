function Get-MspAccessToken {
    <#
    .SYNOPSIS
        Returns a cached or freshly issued delegated access token for one
        customer tenant and one resource.
    .DESCRIPTION
        1. Resolves -TenantId (GUID or verified domain) to a tenant GUID. There
           is no default tenant. Your own partner tenant is only used with
           -PartnerTenant, and passing the partner tenant ID as -TenantId is
           refused so it cannot happen by accident.
        2. Looks up the module-scoped cache (key "<tenantId>|<resource>|<appId>")
           and reuses the entry when Test-MspAccessToken passes (more than 5
           minutes left, same tenant, same resource, required scopes present).
        3. Otherwise reads the technician refresh token from the vault and
           redeems it at the CUSTOMER tenant's v2.0 token endpoint, so the
           technician's GDAP roles in that customer decide what the token can do.
        4. Writes the rotated refresh token back to the vault, then caches the
           new access token as a SecureString.

        Token values never appear in verbose output, warnings or errors.
    .PARAMETER TenantId
        Customer tenant GUID or verified domain. Mandatory unless -PartnerTenant is used.
    .PARAMETER PartnerTenant
        Get a token for your own partner tenant (for example to list GDAP
        relationships or Partner Center calls).
    .PARAMETER Resource
        Resource alias (Graph, Exchange, PartnerCenter, ManagementApi, Defender,
        AzureManagement, TeamsAdmin), resource URI or application ID. Defaults
        to Microsoft Graph. The token is requested with <resource>/.default.
    .PARAMETER Scope
        Specific scopes for one resource instead of .default, for example
        'User.Read.All' or 'https://outlook.office365.com/Exchange.Manage'.
    .PARAMETER UserPrincipalName
        Technician whose refresh token is used. Defaults to the technician
        registered in this session, then TechnicianUpn from the configuration.
    .PARAMETER ForceRefresh
        Skip the cache and redeem the refresh token now.
    .PARAMETER AsPlainText
        Return only the access token string. Use with care.
    .OUTPUTS
        MspGdap.AccessToken (AccessToken is a SecureString), or a string with -AsPlainText.
    .EXAMPLE
        Get-MspAccessToken -TenantId 'contoso.onmicrosoft.com'
    .EXAMPLE
        Get-MspAccessToken -TenantId $customerTenantId -Resource Exchange
    .EXAMPLE
        Get-MspAccessToken -PartnerTenant -Resource PartnerCenter
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    [OutputType('MspGdap.AccessToken')]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant', Position = 0, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = 'Partner')]
        [switch]$PartnerTenant,

        [string]$Resource,

        [string[]]$Scope,

        [string]$UserPrincipalName,

        [switch]$ForceRefresh,

        [switch]$AsPlainText
    )

    process {
        try {
            $config = Get-MspConfigurationInternal -RequireComplete
            $targetTenant = Resolve-MspTargetTenant -TenantId $TenantId -PartnerTenant:$PartnerTenant -Configuration $config
            $request = Resolve-MspTokenRequestScope -Resource $Resource -Scope $Scope
            $upn = Resolve-MspTechnicianUpn -UserPrincipalName $UserPrincipalName -Configuration $config
            $key = Get-MspTokenCacheKey -TenantId $targetTenant -Resource $request.Resource -AppId $config.AppId

            $entry = $null
            $fromCache = $false
            if (-not $ForceRefresh -and $script:MspTokenCache.ContainsKey($key)) {
                $cached = $script:MspTokenCache[$key]
                if ($cached.UserPrincipalName -eq $upn) {
                    $testParams = @{
                        InputObject = $cached
                        TenantId    = $targetTenant
                        Resource    = $request.Resource
                        SkewMinutes = $script:MspConstants.DefaultSkewMinutes
                    }
                    if ($request.ShortScopes.Count -gt 0) { $testParams.RequiredScope = $request.ShortScopes }
                    if (Test-MspAccessToken @testParams) {
                        $entry = $cached
                        $fromCache = $true
                        Write-Verbose "Token cache hit for $key (expires $($cached.ExpiresOn.ToString('u')))."
                    }
                    else {
                        Write-Verbose "Cached token for $key is no longer usable. Requesting a new one."
                    }
                }
                else {
                    Write-Verbose "Cached token for $key belongs to another technician. Requesting a new one."
                }
            }

            if (-not $entry) {
                Write-Verbose "Redeeming the refresh token for $upn at tenant $targetTenant for $($request.Resource)."
                $entry = Invoke-MspRefreshTokenRedemption -Configuration $config -TenantId $targetTenant -Request $request -UserPrincipalName $upn
                $script:MspTokenCache[$key] = $entry
                Write-Verbose "Cached new token for $key (expires $($entry.ExpiresOn.ToString('u')))."
            }

            if ($AsPlainText) {
                ConvertFrom-MspSecureString -SecureString $entry.AccessToken
            }
            else {
                ConvertTo-MspAccessTokenOutput -Entry $entry -FromCache:$fromCache
            }
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}
